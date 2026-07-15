# Output files — what's what

## The main file
**`MAIN - PBSC Overlap Analysis.xlsx`** — start here (named to stand out from
everything else in this folder, which all happen to start with "PBSC"). The
aggregated findings: Summary, By PBSC, Top Connectors, Top Pair Evidence,
Anchors, Manual vs Physical, plus a `Wikipedia Coverage (All 219)` sheet. Every
column that holds Baidu-derived data is prefixed `Baidu ...` so its source is
explicit; `Top Connectors` and `Top Pair Evidence` also carry joined-in
Wikipedia ZH/EN status + office columns for cross-checking. Built by
`scripts/apply_wikipedia_to_main.py`, which combines `scripts/analyze_overlap.py`'s
output with `scripts/enrich_wikipedia.py`'s output.

## Supporting derived data (same underlying Baidu episodes, different shape)
- **`PBSC Overlap - Supporting Network Dataset.xlsx`** + `PBSC visualization
  edges.csv` + `PBSC visualization person-pbsc summary.csv` — an edge-list-shaped
  cut of the same overlap data (one row per person-PBSC-anchor), with seniority
  ranking and confidence heuristics added. Meant as input to an actual network
  graphing tool (Gephi etc.) — nothing currently renders a picture from it.
  Built by `scripts/make_visualization_dataset.py`.

## Validation (not merged into the analysis)
- **`PBSC Wikipedia cross-check.xlsx`** — full detail behind the Wikipedia
  columns joined into the main file: per-person Wikipedia match status, resolved
  page, parsed offices, and a `Baidu Gaps Filled` sheet. This is a second
  opinion sitting *next to* the analysis, not merged into the actual overlap
  scoring — the Top Connectors/Top Pair Evidence rankings in the main file are
  still computed from Baidu data only.

## Narrative write-up
- **`PBSC Career Overlap Findings.docx`** — hand-authored interpretive summary
  (method, per-PBSC cluster readouts, caveats). Not auto-regenerated.

## Raw inputs
Live in `../source/`, not here: `Chinese Leadership Database - Baidu enriched
working copy.xlsx` (the master data — all 219 people, parsed episodes, the
pre-aggregation candidate table) and `PBSC Factionalism - working copy.xlsx`
(the hand-coded manual faction assignments).

## Raw scrape cache and logs
Moved out of this folder into `../scrape_cache/`, since they're hundreds of
fragmented raw files you'd never open directly, not analysis output:
- `baidu_cache/` (339 files) and `wikipedia_cache/` (826 files) — raw cached
  API/HTML responses, so reruns of the scrapers don't hit the network again.
- `baidu_enrichment_report.json` — run log from the Baidu scrape (match rates,
  errors).

## Archive
`archive/` holds superseded versions:
- `PBSC overlap analysis tables - with titles.xlsx` + its `.md` text mirror —
  the pre-Wikipedia, unlabeled-column version of the main file.
- `PBSC overlap analysis tables.xlsx` + `PBSC overlap analysis report.md` — an
  even older pass, before per-PBSC titles were added.
- `PBSC location overlap counts.xlsx` — an earlier, narrower anchor-pair cut,
  superseded by `PBSC Overlap - Supporting Network Dataset.xlsx` two days later
  (same "PBSC Location Summary" idea, more fields).
