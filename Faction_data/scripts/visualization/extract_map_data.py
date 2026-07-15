"""
Extracts province-level career movements for the 7 PBSC members + everyone
in Connections Data.xlsx's Top Connectors sheet, for the China map/timeline
visualization.

Reuses build_episode_pool() from build_unified_connections.py (same pooled
Baidu+Wikipedia episodes already validated there), filters to province-type
anchors only (institutions have no map coordinate), and merges consecutive/
overlapping episodes at the same province into a single continuous stay.

Each non-PBSC person is assigned one "primary PBSC" -- whichever PBSC pairing
has the highest Score in All Candidates (raw) -- used for color-tinting in
the visualization.

Output: map_data.json next to this script, consumed directly by the artifact.
"""

import json
import sys
from pathlib import Path

import openpyxl

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR.parent / "data_collection"))
sys.path.insert(0, str(SCRIPT_DIR.parent / "analysis"))
import enrich_baidu as eb  # noqa: E402
from build_unified_connections import (  # noqa: E402
    SOURCE_WORKBOOK,
    GOOD_WIKI_STATUSES,
    api_get,
    build_episode_pool,
    detect_places_en,
    extract_page,
    load_baidu_episodes_by_name,
    load_wiki_titles_by_name,
    parse_infobox_offices,
    parse_wiki_year,
    rows_as_dicts,
)

OUT_JSON = SCRIPT_DIR / "map_data.json"
CONNECTIONS_XLSX = SCRIPT_DIR.parent.parent / "outputs" / "Connections Data.xlsx"  # scripts/visualization/ -> scripts/ -> Faction_data/

# Fixed 20th Politburo Standing Committee protocol order -- categorical hue
# order must be fixed, never re-sorted by score/cluster size.
PBSC_ORDER = ["习近平", "李强", "赵乐际", "王沪宁", "蔡奇", "丁薛祥", "李希"]
PBSC_COLORS = {
    "习近平": {"light": "#2a78d6", "dark": "#3987e5"},
    "李强":   {"light": "#1baf7a", "dark": "#199e70"},
    "赵乐际": {"light": "#eda100", "dark": "#c98500"},
    "王沪宁": {"light": "#008300", "dark": "#008300"},
    "蔡奇":   {"light": "#4a3aa7", "dark": "#9085e9"},
    "丁薛祥": {"light": "#e34948", "dark": "#e66767"},
    "李希":   {"light": "#e87ba4", "dark": "#d55181"},
}

# Schematic (not cartographically precise) relative positions on a 900x700
# viewBox, laid out west->east and north->south to roughly match real China
# geography -- chosen because self-contained artifacts can't fetch real
# province boundary GeoJSON, and hand-approximated boundary paths risk
# looking distorted. Coordinates are for a labeled-node "dot map," not a
# choropleth.
PROVINCE_COORDS = {
    "新疆": (80, 220), "西藏": (140, 400), "青海": (270, 300), "甘肃": (340, 240),
    "宁夏": (400, 235), "内蒙古": (480, 150), "黑龙江": (700, 70), "吉林": (670, 150),
    "辽宁": (630, 210), "北京": (560, 215), "天津": (580, 230), "河北": (555, 245),
    "山西": (500, 260), "陕西": (430, 290), "山东": (585, 295), "河南": (515, 315),
    "江苏": (605, 340), "安徽": (560, 355), "湖北": (500, 375), "四川": (370, 380),
    "重庆": (425, 395), "贵州": (415, 445), "云南": (330, 470), "湖南": (490, 415),
    "江西": (555, 405), "浙江": (615, 390), "上海": (630, 360), "福建": (595, 445),
    "广西": (460, 485), "广东": (545, 485), "海南": (505, 545), "台湾": (650, 460),
    "香港": (565, 500), "澳门": (548, 505),
}

MIN_YEAR, MAX_YEAR = 1968, 2026

# National-level titles that are unambiguously Beijing-headquartered even
# when the source text never says "Beijing" (e.g. "总书记" never literally
# mentions a location). Baidu bios routinely drop the place once someone
# reaches this tier -- without this, carry-forward logic pins a person at
# their last real posting forever, e.g. leaving Xi Jinping "stuck" in
# Shanghai (his 2007 stop) for the rest of his career, which is wrong from
# the moment he becomes a Politburo Standing Committee member months later.
TOP_LEADERSHIP_TERMS_CN = [
    "总书记", "国家主席", "国家副主席", "中央政治局常委", "中央政治局委员",
    "中央书记处书记", "国务院总理", "国务院副总理", "全国人大常委会委员长",
    "全国人大常委会副委员长", "全国政协主席", "全国政协副主席",
    "中央军事委员会主席", "国家监察委员会主任", "最高人民法院院长", "最高人民检察院检察长",
]
TOP_LEADERSHIP_TERMS_EN = [
    "General Secretary", "President of China", "President of the People's Republic",
    "Vice President of China", "Premier of China", "Vice Premier of China",
    "Politburo Standing Committee", "Member of the Politburo",
    "Chairman of the Central Military Commission", "Chairman of the National People's Congress",
    "Chairman of the Chinese People's Political Consultative",
]


def clean_name(name):
    import re
    return re.sub(r"[（(].*?[）)]", "", str(name or "")).strip()


def province_span_tuples_from_pool(episodes):
    """(anchor, start, end, source, text) tuples for a person's real
    province-anchored episodes from the pooled Baidu+Wikipedia data."""
    spans = []
    for ep in episodes:
        for a in ep["anchors"]:
            if a["anchor"] in PROVINCE_COORDS:
                start, end = a["interval"]
                spans.append((a["anchor"], start, end, ep["source"], ep["text"]))
    return spans


def matches_top_leadership(text, terms):
    """Substring match against the term list, with one deliberate exception:
    "中央书记处书记" (CPC Central Secretariat) also appears, coincidentally,
    inside "共青团中央书记处书记" (Communist Youth League Central Secretariat
    -- a real but different, junior body). Confirmed against this dataset:
    every actual hit is still Beijing-based (CYL Central Committee HQ is also
    Beijing), so excluding it doesn't change any location -- but the match
    should be correct by construction, not by coincidence."""
    for term in terms:
        idx = text.find(term)
        if idx == -1:
            continue
        if term == "中央书记处书记" and text[:idx].endswith("共青团"):
            continue
        return True
    return False


def implicit_beijing_from_baidu(chinese_name, baidu_by_name):
    """Same tuple shape as above, for Baidu episodes with NO usable map
    anchor (no province, or only an excluded/institution anchor like 中央党校)
    whose role text names an unambiguously Beijing-based national title."""
    extra = []
    for r in baidu_by_name.get(chinese_name, []):
        places = eb.split_tags(r.get("Detected Places"))
        if any(p in PROVINCE_COORDS for p in places):
            continue  # already has a real province anchor, leave it alone
        role = r.get("Role Text (CN)") or ""
        if not matches_top_leadership(role, TOP_LEADERSHIP_TERMS_CN):
            continue
        interval = eb.year_interval(r.get("Start Year"), r.get("End Year"))
        if not interval:
            continue
        text = f"{r.get('Start Year')}-{r.get('End Year') or ''} {role} [implied: Beijing]".strip()
        extra.append(("北京", interval[0], interval[1], "Baidu", text))
    return extra


def implicit_beijing_from_wikipedia(chinese_name, wiki_titles):
    """Same idea, scanning Wikipedia infobox offices that have no province
    anchor but do name a top national title (ZH office text or EN office
    text, matched against the matching-language term list)."""
    extra = []
    for lang, terms, detect_fn in (("zh", TOP_LEADERSHIP_TERMS_CN, eb.detect_places), ("en", TOP_LEADERSHIP_TERMS_EN, detect_places_en)):
        status, title = wiki_titles.get((chinese_name, lang), (None, None))
        if status not in GOOD_WIKI_STATUSES or not title:
            continue
        page = extract_page(api_get(lang, title))
        if not page:
            continue
        for o in parse_infobox_offices(page["wikitext"]):
            raw_anchors = eb.split_tags(detect_fn(o["office"])) if lang == "zh" else detect_fn(o["office"])
            if any(a in PROVINCE_COORDS for a in raw_anchors):
                continue
            if not matches_top_leadership(o["office"], terms):
                continue
            start = parse_wiki_year(o["term_start"])
            if start is None:
                continue
            end = parse_wiki_year(o["term_end"])
            interval = eb.year_interval(start, end)
            if not interval:
                continue
            text = f"{o['term_start']}-{o['term_end'] or 'present'} {o['office']} [implied: Beijing]".strip()
            extra.append(("北京", interval[0], interval[1], f"Wikipedia-{lang.upper()}", text))
    return extra


def merge_span_tuples(spans):
    """Collapses (anchor, start, end, source, text) tuples into continuous stays."""
    spans = sorted(spans, key=lambda s: (s[0], s[1]))
    merged = []
    for province, start, end, source, text in spans:
        if merged and merged[-1]["province"] == province and start <= merged[-1]["end"] + 1:
            merged[-1]["end"] = max(merged[-1]["end"], end)
            merged[-1]["sources"].add(source)
        else:
            merged.append({"province": province, "start": start, "end": end, "sources": {source}, "example": text})
    for m in merged:
        m["sources"] = sorted(m["sources"])
    return sorted(merged, key=lambda m: m["start"])


def main():
    src = openpyxl.load_workbook(SOURCE_WORKBOOK, read_only=True, data_only=True)
    cc_rows = rows_as_dicts(src["20th Central Committee"])
    title_by_name = {str(r.get("Chinese Name")).strip(): r.get("Title") for r in cc_rows if r.get("Chinese Name")}
    pinyin_by_name = {str(r.get("Chinese Name")).strip(): str(r.get("Pinyin Name") or "").strip() for r in cc_rows if r.get("Chinese Name")}
    src.close()

    conn = openpyxl.load_workbook(CONNECTIONS_XLSX, read_only=True, data_only=True)
    tc_rows = rows_as_dicts(conn["Top Connectors"])
    connector_names = [clean_name(r["Person"]) for r in tc_rows]

    # Primary PBSC per person = highest-Score pairing in All Candidates (raw).
    raw_rows = rows_as_dicts(conn["All Candidates (raw)"])
    best_pbsc = {}
    for r in raw_rows:
        person = clean_name(r["Person Chinese Name"])
        score = r["Score"] or 0
        if person not in best_pbsc or score > best_pbsc[person][1]:
            best_pbsc[person] = (r["PBSC Chinese Name"], score)
    conn.close()

    baidu_by_name = load_baidu_episodes_by_name()
    wiki_titles = load_wiki_titles_by_name()

    all_people = PBSC_ORDER + [n for n in connector_names if n not in PBSC_ORDER]
    print(f"Extracting movements for {len(all_people)} people ({len(PBSC_ORDER)} PBSC + {len(all_people) - len(PBSC_ORDER)} connected)...")

    people_out = []
    for name in all_people:
        pool = build_episode_pool(name, baidu_by_name, wiki_titles)
        span_tuples = province_span_tuples_from_pool(pool)
        span_tuples += implicit_beijing_from_baidu(name, baidu_by_name)
        span_tuples += implicit_beijing_from_wikipedia(name, wiki_titles)
        spans = merge_span_tuples(span_tuples)
        if not spans and name not in PBSC_ORDER:
            continue  # no province-level movement data at all -- nothing to plot
        is_pbsc = name in PBSC_ORDER
        primary_pbsc = name if is_pbsc else best_pbsc.get(name, (None, 0))[0]
        people_out.append({
            "name_cn": name,
            "name_en": pinyin_by_name.get(name, ""),
            "title": title_by_name.get(name, ""),
            "is_pbsc": is_pbsc,
            "primary_pbsc": primary_pbsc,
            "spans": [
                {"province": s["province"], "start": s["start"], "end": s["end"], "sources": s["sources"], "example": s["example"][:160]}
                for s in spans
            ],
        })

    with_data = sum(1 for p in people_out if p["spans"])
    print(f"{len(people_out)} people included, {with_data} with at least one province-level span.")

    output = {
        "min_year": MIN_YEAR,
        "max_year": MAX_YEAR,
        "pbsc_order": PBSC_ORDER,
        "pbsc_colors": PBSC_COLORS,
        "province_coords": PROVINCE_COORDS,
        "people": people_out,
    }
    OUT_JSON.write_text(json.dumps(output, ensure_ascii=False, indent=None), encoding="utf-8")
    print(f"Wrote {OUT_JSON} ({OUT_JSON.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
