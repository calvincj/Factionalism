import re
import zipfile
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl


ROOT = Path(r"C:\Users\AMatthias\Documents\Codex\2026-05-12\i-want-to-make-database-that")
TABLES = ROOT / "PBSC overlap analysis tables.xlsx"
OUT = ROOT / "PBSC Career Overlap Findings.docx"


def rows_as_dicts(ws):
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(v or "").strip() for v in rows[0]]
    out = []
    for values in rows[1:]:
        item = {}
        for idx, header in enumerate(headers):
            item[header] = values[idx] if idx < len(values) else None
        if any(v not in (None, "") for v in item.values()):
            out.append(item)
    return out


def text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def clean(value):
    value = text(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def run_xml(value, bold=False, italic=False):
    value = text(value)
    props = ""
    if bold or italic:
        parts = []
        if bold:
            parts.append("<w:b/>")
        if italic:
            parts.append("<w:i/>")
        props = "<w:rPr>" + "".join(parts) + "</w:rPr>"
    pieces = value.split("\n")
    runs = []
    for idx, piece in enumerate(pieces):
        if idx:
            runs.append("<w:r><w:br/></w:r>")
        runs.append(f'<w:r>{props}<w:t xml:space="preserve">{escape(piece)}</w:t></w:r>')
    return "".join(runs)


def para(value="", style=None, bold=False, italic=False):
    ppr = f"<w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>" if style else ""
    return f"<w:p>{ppr}{run_xml(value, bold=bold, italic=italic)}</w:p>"


def bullet(value):
    return para("- " + text(value), style="ListParagraph")


def cell(value, shade=None, bold=False):
    shd = f'<w:shd w:fill="{shade}"/>' if shade else ""
    return (
        "<w:tc><w:tcPr>"
        "<w:tcW w:w=\"2400\" w:type=\"dxa\"/>"
        f"{shd}</w:tcPr>"
        f"{para(text(value), bold=bold)}"
        "</w:tc>"
    )


def table(headers, data_rows):
    tbl_pr = (
        "<w:tblPr>"
        "<w:tblStyle w:val=\"TableGrid\"/>"
        "<w:tblW w:w=\"0\" w:type=\"auto\"/>"
        "<w:tblBorders>"
        "<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"B7B7B7\"/>"
        "<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"B7B7B7\"/>"
        "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"B7B7B7\"/>"
        "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"B7B7B7\"/>"
        "<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D9D9D9\"/>"
        "<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D9D9D9\"/>"
        "</w:tblBorders>"
        "</w:tblPr>"
    )
    rows = ["<w:tr>" + "".join(cell(h, shade="D9EAF7", bold=True) for h in headers) + "</w:tr>"]
    for row in data_rows:
        rows.append("<w:tr>" + "".join(cell(v) for v in row) + "</w:tr>")
    return "<w:tbl>" + tbl_pr + "".join(rows) + "</w:tbl>"


def styles_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="180"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="34"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/>
    <w:rPr><w:i/><w:color w:val="666666"/><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="360" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="1F4E79"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="2F5597"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="360"/></w:pPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:tblPr><w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="B7B7B7"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="B7B7B7"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="B7B7B7"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="B7B7B7"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>
    </w:tblBorders></w:tblPr>
  </w:style>
</w:styles>"""


def content_types_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""


def rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def document_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def core_xml():
    today = date.today().isoformat()
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>PBSC Career Overlap Findings</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{today}T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{today}T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""


def app_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Word</Application>
</Properties>"""


def document_xml(body):
    sect = (
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1080" w:right="900" w:bottom="1080" w:left="900" w:header="720" w:footer="720" w:gutter="0"/>'
        "</w:sectPr>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'mc:Ignorable="w14 wp14"><w:body>'
        + "".join(body)
        + sect
        + "</w:body></w:document>"
    )


def main():
    wb = openpyxl.load_workbook(TABLES, read_only=True, data_only=True)
    summary = {r["Metric"]: r["Value"] for r in rows_as_dicts(wb["Summary"])}
    by_pbsc = rows_as_dicts(wb["By PBSC"])
    top_connectors = rows_as_dicts(wb["Top Connectors"])
    top_pairs = rows_as_dicts(wb["Top Pair Evidence"])
    anchors = rows_as_dicts(wb["Anchors"])
    manual = rows_as_dicts(wb["Manual vs Physical"])

    manual_counts = {}
    for row in manual:
        manual_counts[row["Category"]] = manual_counts.get(row["Category"], 0) + 1
    manual_only = [r for r in manual if r["Category"] == "manual_only"]

    pbsc_order = ["Li Qiang", "Ding Xuexiang", "Zhao Leji", "Xi Jinping", "Cai Qi", "Li Xi", "Wang Huning"]
    by_pbsc_sorted = sorted(by_pbsc, key=lambda r: pbsc_order.index(r["PBSC"]) if r["PBSC"] in pbsc_order else 99)

    body = []
    body.append(para("PBSC Career Overlap Findings", style="Title"))
    body.append(para("Strict physical-overlap analysis of 20th Central Committee career histories against the 20th PBSC", style="Subtitle"))
    body.append(para("Prepared May 18, 2026. Central Party School overlaps are excluded from the mechanical overlap counts. Manual faction-workbook labels are preserved separately.", italic=True))

    body.append(para("Executive Summary", style="Heading1"))
    body.append(bullet(f"The dataset includes {summary.get('central_committee_rows')} Central Committee rows and {summary.get('baidu_episode_rows')} parsed career episodes."))
    body.append(bullet(f"After excluding Central Party School as an overlap anchor, the strict candidate set contains {summary.get('candidate_rows_total')} rows, including {summary.get('physical_candidate_rows')} physical-overlap rows across {summary.get('unique_people_with_physical_overlap')} non-PBSC individuals."))
    body.append(bullet(f"There are {summary.get('manual_assignments')} manual PBSC faction assignments from the separate faction workbook; {manual_counts.get('manual_and_physical', 0)} also have detected physical overlap, while {manual_counts.get('manual_only', 0)} remain manual-only."))
    body.append(bullet("The strongest clusters are Li Qiang-Zhejiang, Ding Xuexiang-Shanghai, Zhao Leji-Qinghai/Shaanxi, Xi Jinping-Fujian/Zhejiang/Shanghai/Tsinghua, Cai Qi-Zhejiang/Fujian/Beijing, and Li Xi-Guangdong/Liaoning/Shaanxi/Gansu/Shanghai. Wang Huning has the sparsest physical-overlap network."))

    body.append(para("Method", style="Heading1"))
    body.append(para("The analysis uses Baidu-derived career timelines parsed into start year, end year, role text, and detected physical anchors. A candidate overlap requires overlapping years and the same specific place or named institution/organization. Broad tags such as provincial, academia, party_center, SOE, PLA, and State Council are retained only as context. Central Party School was removed as an overlap anchor at your request."))
    body.append(para("The resulting rows should be read as contact-opportunity evidence rather than direct proof of factional alignment. Manual labels from the faction workbook are marked separately so they can be compared with the mechanical results."))

    body.append(para("PBSC-Level Findings", style="Heading1"))
    body.append(table(
        ["PBSC", "Physical People", "Manual Labels", "Physical Pairs", "Top Places", "Top Institutions"],
        [
            [
                r["PBSC"],
                r["Physical People"],
                r["Manual People"],
                r["Physical Pairs"],
                clean(r["Top Places"]),
                clean(r["Top Institutions"]),
            ]
            for r in by_pbsc_sorted
        ],
    ))

    body.append(para("Cluster Readout", style="Heading1"))
    cluster_notes = [
        ("Li Qiang", "The strongest identifiable network is Zhejiang, with Shanghai and Jiangsu as secondary anchors. The top overlaps are Zhao Yide, Ying Yong, Lou Yangsheng, Chen Min'er, Chen Yixin, and Shen Yueyue."),
        ("Ding Xuexiang", "The pattern is overwhelmingly Shanghai. Wang Wentao and Ye Jianchun are also manual faction-workbook matches; Xu Lin, Wang Xi, Shen Xiaoming, Yin Hong, and Tang Dengjie have strong physical Shanghai overlap."),
        ("Zhao Leji", "Qinghai is the clearest substantive anchor. Qu Qingshan, Wang Yubo, Yan Jinhai, and Cheng Lihua are especially strong. Wang Lixia remains a manual and physical Shaanxi-linked case."),
        ("Xi Jinping", "After removing Central Party School, the network narrows to Fujian, Zhejiang, Shanghai, and Tsinghua. Wang Xiaohong remains a strong Fujian-linked case and a manual workbook match."),
        ("Cai Qi", "The network is concentrated in Zhejiang, Beijing, and Fujian. It overlaps heavily with broader Xi and Li Qiang networks, so manual interpretation is still needed to decide whether a connection is more Cai-specific or part of a wider Zhejiang/Fujian career path."),
        ("Li Xi", "The pattern is more dispersed: Guangdong, Liaoning, Shaanxi, Gansu, and Shanghai. Manual matches Wang Weizhong and Wang Zhengpu are consistent with physical overlap."),
        ("Wang Huning", "The physical network is small: Wang Xiaohui through the Central Policy Research Office and Wang Wentao through Fudan. Yin Li remains a manual-only note from the faction workbook."),
    ]
    for title, note in cluster_notes:
        body.append(para(title, style="Heading2"))
        body.append(para(note))

    body.append(para("Strongest Cross-PBSC Connectors", style="Heading1"))
    body.append(table(
        ["Person", "Title", "PBSC Count", "Score", "Physical Pairs", "PBSCs", "Top Anchors"],
        [
            [
                f"{r['Person']} / {r['Pinyin']}",
                clean(r.get("Title")),
                r["PBSC Count"],
                r["Score"],
                r["Physical Pairs"],
                clean(r["PBSCs"]),
                clean(r["Top Anchors"]),
            ]
            for r in top_connectors[:15]
        ],
    ))

    body.append(para("Strongest Pair-Level Evidence", style="Heading1"))
    body.append(table(
        ["Person -> PBSC", "Title", "Score", "Pairs", "Anchors", "Years"],
        [
            [
                f"{r['Person']} / {r['Pinyin']} -> {r['PBSC']}",
                clean(r.get("Title")),
                r["Score"],
                r["Pairs"],
                clean(r["Anchors"]),
                f"{r['Start']}-{r['End']}",
            ]
            for r in top_pairs[:15]
        ],
    ))

    body.append(para("Anchor Concentration", style="Heading1"))
    body.append(table(
        ["Anchor", "Type", "Candidate Rows", "Physical Pair Count"],
        [[r["Anchor"], r["Type"], r["Candidate Rows"], r["Physical Pair Count"]] for r in anchors[:15]],
    ))

    body.append(para("Manual Labels Without Detected Physical Overlap", style="Heading1"))
    if manual_only:
        for r in manual_only:
            body.append(bullet(f"{r['Person']} -> {r['PBSC']}"))
    else:
        body.append(para("None."))

    body.append(para("Caveats And Next Steps", style="Heading1"))
    caveats = [
        "A shared province or institution is evidence of possible career contact, not proof of patronage or factional loyalty.",
        "Some overlaps use year-level dates, so month-level non-overlap may remain hidden where the source timeline is coarse.",
        "Twenty-six biographies still need non-Baidu fallback research before this should be treated as complete.",
        "The strongest next refinement would be separating same-province leadership overlap from same-campus/institution overlap, then adding a confidence label for each pair.",
    ]
    for item in caveats:
        body.append(bullet(item))

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml())
        zf.writestr("_rels/.rels", rels_xml())
        zf.writestr("word/_rels/document.xml.rels", document_rels_xml())
        zf.writestr("word/document.xml", document_xml(body))
        zf.writestr("word/styles.xml", styles_xml())
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("docProps/app.xml", app_xml())

    print(OUT)


if __name__ == "__main__":
    main()
