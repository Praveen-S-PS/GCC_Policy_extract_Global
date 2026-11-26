# Annual Graduates Pool

Estimate the annual pool of **STEM** and **non-STEM** graduates for any city in the world, and export per-city Excel workbooks plus a single consolidated deliverable.

Give the tool a list of cities (and optionally the colleges in them) and it produces, for each city, a Low / Mid / High estimate of graduates per year split into STEM and non-STEM, every figure traceable back to its source.

## How it works

Each city total is built from two layers:

1. **Known institutions (base).** For every college listed for a city, the tool finds a graduate figure using a five tier fallback, best first:
   1. NIRF "Sanctioned (Approved) Intake" tables (primary for India)
   2. A direct convocation / graduate count from the source page
   3. Total enrolment from Wikipedia / Wikidata (`P2196`), divided by course duration and a completion factor
   4. A looser enrolment or intake keyword figure from the source
   5. Otherwise the cell is marked `NA`

2. **Residual estimate.** Cities larger than a threshold get an estimate for the colleges that are not individually listed, so a large city is not undercounted to just its few famous universities. The residual is driven by the city population:

   ```
   total_institutions = population / 1,000,000 x INSTITUTION_DENSITY_PER_M
   residual = (total_institutions - listed) x average graduates per institution
   ```

   Cities at or below the threshold use their listed colleges only, so small towns are not inflated by phantom institutions.

**Populations** come first from a curated metro table (`input/metro_population.json`) so figures are stable between runs, then fall back to a country verified Wikidata lookup for any city not in the table.

**STEM split** uses curated per city overrides where available, otherwise a field based heuristic (engineering, medical, and technical fields weighted toward STEM; arts, commerce, and management toward non-STEM).

Reference cities that were calibrated against a known data set are pinned to exact figures; every other city is produced automatically by the formula above.

## Requirements

* Python 3.11 or newer
* Packages: `openpyxl`, `requests`, `urllib3`, `beautifulsoup4`, `pypdf`

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl requests urllib3 beautifulsoup4 pypdf
```

## Quick start

```bash
# Process every city in the input list and build the deliverable
python3 generate_graduates.py

# Preview the cities that would be processed, no scraping
python3 generate_graduates.py --list

# Process a single country or a single city
python3 generate_graduates.py --country "India"
python3 generate_graduates.py --city "Chennai"
```

An internet connection is needed on the first run. Successful and failed fetches are cached in `.cache/`, so later runs are much faster.

## Inputs

All inputs live in `input/`.

| File | Columns | Purpose |
|------|---------|---------|
| `countries.xlsx` | `Country`, `City` | The master city list and its order |
| `universities.xlsx` | `Country`, `City`, `College Name`, `Source List` | Colleges per city. `Source List` (a URL) is optional; blank rows are resolved by name through Wikipedia / Wikidata |
| `metro_population.json` | `"country\|city": population` | Curated, stable city populations |
| `template.xlsx` | | Per city workbook template |
| `Final_Template.xlsx` | | Header only template for the consolidated deliverable |
| `base_overrides.json`, `city_profiles_extra.json`, `stem_overrides.json` | | Curated overrides for calibrated reference cities |

To add a city, add a row to `countries.xlsx`. List its colleges in `universities.xlsx` for precision, or leave them out and the tool will discover institutions automatically from Wikidata. For a stable population, add an entry to `metro_population.json`.

## Outputs

Every run writes to a dated snapshot folder, `output/data - {YYYY-MM-DD}/`:

* `{Country}/{City} - graduates - {date}.xlsx` — a detailed per city workbook (listed institutions, residual estimate, city total, STEM share)
* `Annual Graduates Pool - {date}.xlsx` — the consolidated deliverable, one row per city with Population and STEM / non-STEM Low, Mid, High

## Command line options

| Flag | Effect |
|------|--------|
| `--country NAME` | Only process this country |
| `--city NAME` | Only process this city |
| `--list` | List the cities in the input and exit |
| `--check-population` | Report which cities have a curated (stable) population versus rely on the live lookup, then exit |
| `--summary-only` | Rebuild the deliverable from existing per city workbooks, no scraping (needs `--output-dir` pointed at a dated folder) |
| `--no-summary` | Do not write the consolidated deliverable |
| `--no-population` | Skip the Population lookup |
| `--offline` | Use the cache only, do not hit the network |
| `--no-wikidata` | Disable the Wikipedia / Wikidata enrolment fallback |
| `--enrich-sources` | Write a copy of the college list with resolved Wikipedia and official site links |
| `--merge-urls FILE` | Merge an external `Institution, Source URLs` file into the college list |
| `--workers N` | Parallel fetch workers (default 8; politeness is enforced per host) |
| `--date STR` | Date string used in filenames (default: today, ISO) |

Run `python3 generate_graduates.py --help` for the full list.

## Project structure

```
Graduates_pool/
  generate_graduates.py    # command line entry point (argparse)
  gradpool_core.py         # config, networking, scraping, parsing, lookups
  gradpool_excel.py        # all Excel reading and writing
  input/                   # all inputs (see table above)
  output/                  # dated result snapshots
  .cache/                  # fetch cache (safe to delete)
```

Dependency order is `gradpool_core` then `gradpool_excel` then `generate_graduates`.

## Calibration

Formula constants live near the top of `gradpool_core.py` and are tunable:

| Constant | Default | Meaning |
|----------|---------|---------|
| `INSTITUTION_DENSITY_PER_M` | 36 | Higher education institutions per million people |
| `RESIDUAL_AVG_LOW/MID/HIGH` | 55 / 80 / 110 | Average graduates per residual institution |
| `RESIDUAL_STEM_SHARE` | 0.30 | STEM share applied to the residual |
| `BASE_CALIBRATION` | 0.41 | Scales scraped base figures toward the reference level |
| `STEM_CALIBRATION` | 0.85 | Damps the field heuristic STEM share |

## Accuracy

These numbers are best effort estimates built from public data (NIRF, Wikipedia, Wikidata, and curated reference sets). Automated figures can be wrong where a source is missing, blocked, or ambiguous, and the residual estimate assumes an average institution density that will not hold for every city. Every value in the per city workbooks carries a note recording the figure, the method, and the source.

**Please verify the figures against official sources before sharing them externally.**
