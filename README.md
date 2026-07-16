# Factionalism

This is my project on CCP elite factionalism. The main
question I'm trying to answer: how connected are the members of the 20th
Politburo Standing Committee (PBSC) to other Chinese officials, based on
where their careers overlapped over the decades? If two people worked in the
same province or institution at the same time, that's a potential factional
tie, and this project tries to actually measure that instead of just
eyeballing bios.

## What's in here

- **`Faction_data/`** is where basically everything lives: the raw source
  workbooks, the scripts that scrape and analyze them, and the finished
  output files. Start with `Faction_data/outputs/README.md` if you just want
  the results, or `Faction_data/scripts/README.md` if you want to know how
  the pipeline actually works.
- **`References/`** has outside source material I pulled in for context.
- **`Writing/`** has the draft paper and notes that use this data.

## The short version of the pipeline

1. Pull career bios for every relevant official from Baidu Baike and
   Wikipedia (Chinese and English).
2. Parse those bios into "career episodes": a place or institution, plus the
   years someone was there.
3. Match people up by shared place/institution and overlapping years. Each
   match is a "pair," and pairs get pooled into a connection score per
   person relative to each PBSC member.
4. Everything lands in `Faction_data/outputs/Connections Data.xlsx`, and also
   gets turned into an interactive map (`pbsc-connections-map.html`)
   that shows where everyone was posted, year by year, from 1968 to 2026.

## Regenerating stuff

The `Faction_data/scrape_cache/` folder (raw scraped HTML/JSON) is not
checked into git since it's huge and fully regenerable. If you clone this
repo fresh and want to rerun anything from scratch, check
`Faction_data/scripts/README.md` for which script to run and in what order.
# Factionalism
