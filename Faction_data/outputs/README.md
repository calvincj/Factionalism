# Output files

The four data files below, plus one visualization.

1. **`Baidu Data - All People.xlsx`** — every one of the 219 people's Baidu
   match status/URL/summary (`People` sheet), plus all 3,557 parsed career
   episodes (`Career Episodes` sheet). This is "the Baidu data of everyone,"
   pulled out of the much bigger master workbook in `../source/` so you don't
   have to dig through 12 sheets to find it.

2. **`Wikipedia Data - All People.xlsx`** — every person's Chinese + English
   Wikipedia match status, resolved page, parsed infobox offices, and extract
   text. Independent of Baidu — nothing here has been cross-checked yet.

3. **`Merged Data - All People.xlsx`** — one row per person, Baidu and
   Wikipedia side by side, plus a `Cross-Check Flag` column
   (`gap_filled_by_wikipedia` / `both_sources_have_data_review_for_consistency`
   / `baidu_only` / `neither_source_has_structured_data`). This is where you
   go to see, per person, what each source says and whether they agree.

4. **`Connections Data.xlsx`** — the actual factional-overlap analysis:
   Summary, By PBSC, Top Connectors, Top Pair Evidence, Anchors, Manual vs
   Physical. Built by `scripts/analysis/build_unified_connections.py`, which pools each
   person's Baidu episodes and Wikipedia (ZH+EN) offices into one combined set
   of career evidence per person before doing the anchor+interval-overlap
   matching — so a "pair" can come from either source, and no column
   distinguishes which. The `Evidence` text on `Top Pair Evidence` keeps a
   short `[Baidu]`/`[Wikipedia-ZH]`/`[Wikipedia-EN]` tag per line so you can
   still trace a specific claim back to its source if needed.

## Visualization
- **`PBSC Connections Tracking Map.html`** — open directly in a browser (not a Claude
  Artifact, a plain local file). A real China province map (actual GeoJSON
  boundaries, embedded, from the standard Aliyun DataV Geo Atlas source) with
  a year timeline (1968–2026, play button included) showing where the 7 PBSC
  members and 84 connected officials were posted each year. PBSC members are
  colored per the fixed protocol-order palette and always shown (carried
  forward through data gaps); connected officials are neutral grey dots.
  Requires internet on first load only to fetch the ECharts library from a
  CDN — the map data and China boundaries are fully embedded in the file.
  Built by `scripts/visualization/extract_map_data.py` (province-level movement extraction)
  plus a one-off assembly step that embeds the GeoJSON and this data into the
  HTML template.

## Everything else
Moved to `archive/`: an older combined MAIN file, a standalone Wikipedia
cross-check file, the network/visualization dataset + CSVs, the narrative
docx write-up, an even older location-overlap-counts pass, and superseded
versions of the analysis tables. None of it is gone, just out of the way.

## Not in this folder
- `../source/` — the raw master database (`Chinese Leadership Database -
  Baidu enriched working copy.xlsx`, 12 sheets covering multiple Central
  Committee generations, of which "Baidu Career Episodes" and "20th Central
  Committee" are what file 1 above was extracted from) and the hand-coded
  manual faction workbook. This is the actual source of truth the 4 files
  above are derived from — keep it, don't archive it.
- `../scrape_cache/` — raw HTML/JSON scrape cache (hundreds of files), not
  analysis data.
- `../scripts/` — the code that builds all of the above (see `scripts/README.md` for what's where).
