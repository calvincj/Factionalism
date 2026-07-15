# Scripts

Five folders, organized by what each script actually does, not by when it
was written.

## `data_collection/`
Fetches and cleans raw biographical data. Nothing here reads another
script's output — these are the two original sources.
- `enrich_baidu.py` — scrapes Baidu Baike, parses career timelines into
  episodes (place/institution anchors, years, role text). Slow, hits the
  network (or the local `baidu_cache/`).
- `enrich_wikipedia.py` — resolves each person to a Chinese + English
  Wikipedia page, parses infobox office/term data. Imports from
  `enrich_baidu.py` (same folder).

## `analysis/`
The actual factional-connections pipeline — reads `data_collection/`'s
output, produces `outputs/Connections Data.xlsx`.
- `build_unified_connections.py` — the full rebuild: pools Baidu + Wikipedia
  episodes per person, runs the anchor+interval-overlap matching, writes
  every sheet (Summary, By PBSC, Top Connectors, Top Pair Evidence, Anchors,
  Manual vs Bot, All Candidates (raw), Data Dictionary). Run this after
  changing source data.
- `sync_manual_overrides.py` — the fast path: reads a hand-edited `Manual
  PBSCs` column in `Top Connectors` and recomputes everything downstream
  (Score, Manual vs Bot, etc.) without re-matching from scratch. Run this
  after manually editing that one column.

## `visualization/`
- `extract_map_data.py` — pulls province-level career movements for the 7
  PBSC + everyone in `Connections Data.xlsx`'s Top Connectors, for the China
  map. Imports from both `data_collection/` and `analysis/`. Writes
  `map_data.json` (same folder) — the actual HTML map
  (`outputs/pbsc-connections-map.html`) embeds this plus a GeoJSON
  China boundary file via a one-off assembly step, not itself checked in
  here as a script.

## `exports/`
- `consolidate_outputs.py` — builds the 4 "All People" xlsx exports
  (`Baidu Data`, `Wikipedia Data`, `Merged Data`, `Connections Data`, all in
  `outputs/`). **Currently stale** — it still points at two files
  (`MAIN - PBSC Overlap Analysis.xlsx`, `PBSC Wikipedia cross-check.xlsx`)
  that were archived once `build_unified_connections.py` took over producing
  `Connections Data.xlsx` directly. Needs a rewrite before it'll run again,
  not just a path fix.

## `legacy/`
Superseded scripts, kept for reference. Not part of the active pipeline, not
guaranteed to still run (some import from `data_collection/` using the old
flat-folder assumption and would need path fixes first). Each one's output
already lives in `outputs/archive/`:
- `analyze_overlap.py` — the original Baidu-only analysis, before Wikipedia
  was pooled in.
- `apply_wikipedia_to_main.py` — an intermediate step from before
  `build_unified_connections.py` existed, merging Wikipedia into the old
  MAIN file format.
- `make_location_overlap_counts.py` — an early, narrower anchor-pair cut.
- `make_visualization_dataset.py` — the old network/edge-list dataset,
  superseded by the actual working China map.
- `make_word_findings.py` — generates the narrative `.docx` write-up.

## Cross-script imports
Scripts in different folders that import each other add the needed sibling
folder to `sys.path` explicitly (e.g. `analysis/build_unified_connections.py`
adds `../data_collection`) — look at the top of each file for exactly which
folders it needs. Every active script's own `SOURCE_WORKBOOK`/`OUTPUTS`/etc.
path constants account for now living one level deeper than the old flat
`scripts/` layout (`SCRIPT_DIR.parent.parent` to reach `Faction_data/`,
not `SCRIPT_DIR.parent`).
