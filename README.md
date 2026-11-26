# GCC Incentives Tracker

A Python pipeline that reads official government policy PDFs and web pages, picks out the statements that matter for Global Capability Centres (GCCs), and writes them to structured Excel workbooks. Each row is one incentive statement with its scope (country, state, city), validity dates and a GCC relevance score.

Coverage:

* **India:** 8 priority states.
* **State level targets:** United States (50 states), United Kingdom (4 nations), Germany (16 states), France (13 regions) and Japan (47 prefectures).
* **Global:** a country by country sheet for about 330 countries from `input/template.xlsx`, of which about 110 have configured sources.

The repository also holds a separate, older news digest pipeline (see the last section).

## What a run produces

Everything is written under `output/`, with the run date in the file name.

* **Flat workbook** in `output/new_template/<Target>/`. One row per incentive statement with the 12 columns described below. For the Global target it has a `Global` sheet (rows with a GCC signal) and an `Other - Global` sheet. This is the file to analyse and share.
* **Wide audit workbook** in `output/old_template/<Target>/`. One sheet per jurisdiction, with 6 incentive types by 3 scope levels, plus a `gcc_sources_*` workbook listing every source and its status. Skip both with `--no-old-template`.

Examples of file names: `new_template_GCC_policy_2026-09-23.xlsx`, `new_template_GCC_policy_subset_2026-10-08.xlsx` (a `--countries` run), `new_template_india_2026-10-06.xlsx`, `new_template_germany_states_policy_2026-08-25.xlsx`.

## Quick start

```bash
git clone <repository-url>
cd GCC_News_incentive_tracker
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# optional: translate non English sources into English
pip install argostranslate langdetect

# optional: LLM paraphrasing (install Ollama from https://ollama.com first)
ollama pull llama3.2:3b

# a small first run
python3 country_extractor.py --global --countries "Brazil,Spain"
```

Requirements:

* Python 3.10 or newer (developed on 3.14).
* Required packages: `requests`, `pdfplumber`, `openpyxl`, `beautifulsoup4` (all in `requirements.txt`).
* Optional: `argostranslate` and `langdetect` for translation (language packs download on first use, about 200 MB per language pair).
* Optional: Ollama with `llama3.2:3b` for paraphrasing, and `ocrmypdf` with Tesseract for scanned PDFs.

If an optional piece is missing, the pipeline prints a warning and skips that step. Without Ollama, run with `--nopara`.

## Running it

`country_extractor.py` is the single entry point. With no flags it runs every state level target and the Global target, one after another.

```bash
python3 country_extractor.py                        # everything
python3 country_extractor.py --germany --france     # chosen targets
python3 country_extractor.py --india                # India's 8 states only
python3 country_extractor.py --global               # Global sheet plus India
python3 country_extractor.py --global --countries "Brazil,Spain,Italy"   # a subset of countries
```

Target flags: `--usa`, `--uk`, `--germany`, `--france`, `--japan`, `--global`, `--india`.

Other options:

* `--nopara` skips the Ollama paraphrasing step (faster, raw extracted text).
* `--notranslate` skips translation.
* `--refresh` downloads every PDF again, ignoring the local cache.
* `--model NAME` uses another Ollama model (default `llama3.2:3b`).
* `--min-len N` only paraphrases bullets of at least N characters (default 0).
* `--no-old-template` writes only the flat workbook.
* `--countries "A,B,C"` (Global only) rebuilds just those countries and only fetches their sources.

### Long runs

A full run can take hours, mostly because of paraphrasing. Two options make long jobs safe to interrupt:

```bash
python3 country_extractor.py --global --countries "Brazil,Spain,Italy" --resume --max-minutes 8
```

* `--max-minutes N` stops cleanly after about N minutes, at the next paraphrase chunk. The exit code is 3, which means "stopped at the limit, run the same command again".
* `--resume` keeps the fetched and translated text in `output/.resume`, so the next run does not fetch again. The saved text is deleted when the run completes, and anything older than 24 hours is ignored.
* The paraphrase cache (`output/.paraphrase_cache.json`) is saved every 8 items and written atomically, so progress survives a crash. Do not run two jobs at the same time, because they share this cache.
* For big jobs, run slices of about 5 to 10 countries. Each Global subset run writes the same file name for the day, so copy the output away after each slice.

## How the pipeline works

1. **Fetch.** PDFs are downloaded to `sources/` with a freshness check. Downloads are atomic and verified against the server's size, and a cached PDF that is cut off is downloaded again. HTML pages are fetched with a one hop crawl, HTTP 5xx errors are retried, and linked PDFs are read as well.
2. **Read.** PDFs are read with `pdfplumber`. Clearly two column pages are read one column at a time. Scanned PDFs use OCR if `ocrmypdf` is installed.
3. **Translate.** Non English text is translated with Argos Translate before classification, because the keyword lists are English.
4. **Bucket.** `extractor/policy_bucketer.py` sorts paragraphs into 6 incentive types (Financial Incentives, Infrastructure, Talent, Regulatory and Non Financial Support, Innovation and R&D, Others) and 3 scope levels. Paragraphs under 150 characters are dropped unless they are 50 to 149 characters and state a figure (a percentage, amount, duration or thousands figure). Blocks over 3000 characters are split, not dropped. Marketing style sentences, definitions and similar text are skipped.
5. **Paraphrase.** Each statement is rewritten as a short English summary by a local Ollama model and cached by a hash of the model and input.
6. **Build rows.** Cities named in a statement become one row per city. Validity dates come from the statement text or the `Validity` column of the sources file. A row is dropped as expired only if its own text or the `Validity` column says so, never because of a date found elsewhere on the source page.
7. **Clean up.** Text about a different country than the one it is filed under is removed, and near duplicate rows are merged within the same location and section.
8. **Write.** The flat and wide workbooks are saved under `output/`.

## Sources

Sources are plain Excel files in `input/sources/`, so no code change is needed to add or remove one.

* `sources.xlsx` covers India's states and the Global countries.
* `sources_united_states.xlsx`, `sources_united_kingdom.xlsx`, `sources_germany.xlsx`, `sources_france.xlsx` and `sources_japan.xlsx` cover the state level targets.

Columns: `Active` (yes or no), `Region`, `Jurisdiction`, `Filename`, `Validity`, `Source URL`, `Referer` (optional), `Notes`.

Rules worth knowing:

* A PDF URL needs a `Filename` (where the downloaded copy is saved). A PDF row without one is ignored, and the loader prints a warning.
* A row without a `Filename` is treated as an HTML page and crawled one hop. Avoid pages that link to large document libraries, because the crawl will pull them in and flood the country with low value rows.
* Set `Active` to no instead of deleting a row, and record why in `Notes`.
* Every run prints a warning for each source that returned nothing, with the reason (HTTP 403 or 404, SSL error, timeout, bot challenge page, a page that is JavaScript only, or a PDF with no extractable text). Read these after each run. A blocked site is not worked around; find another official page instead.

## Flat workbook columns

* `section_name`: one of the 6 incentive types.
* `scope_type`: Country, State, Tier or City.
* `country_name`, `state_name`, `city_name`: where the statement applies.
* `tier_scope`: Tier 1 or Tier 2 & 3 when the statement is tier specific.
* `description`: the paraphrased (or raw) statement.
* `validity_start_date`, `validity_start_precision`, `validity_end_date`, `validity_end_precision`: validity dates and how precise they are.
* `gcc_relevance`: the score described next.

## GCC relevance score

`gcc_relevance` is a keyword score computed on each row's text by `compute_gcc_relevance()` in `new_template.py`.

* Strong terms add 2: global capability centre, shared services, BPO, KPO, ITES, R&D centre, regional headquarters, delivery centre, back office and similar.
* Medium terms add 1: IP or patent box, R&D tax relief, technology park, innovation hub, centre of excellence.
* Weak terms add 0.5 (innovation, software, outsourcing, offshore), and the total is rounded, so a weak term on its own scores 0.
* India, Mexico and Poland get 1 extra point, because they have GCC specific policies.

Rows scoring 1 or more go to the `Global` sheet, the rest to `Other - Global`. The score is only a label and never removes a row. Most countries have general investment promotion pages, so few rows reach a signal. Adding GCC specific sources for a country is the way to raise it.

## Data quality and known limitations

This pipeline produces a draft for human review. Check the output for accuracy before sharing it externally.

* The paraphrase model can change numbers or drop details from long passages. Verify figures against the source.
* Row counts can differ between runs because live websites change, block automated access or time out. Compare with the previous output before trusting a drop.
* Slide style PDFs (for example some consultancy playbooks) come out jumbled in both reading modes.
* A language without an Argos pack stays untranslated, and the English keyword lists then match less.
* Image only PDFs need OCR, otherwise they give no text.
* HTML landing pages can mix marketing text with policy detail and may need light editing.

## Project layout

```
GCC_News_incentive_tracker/
├── country_extractor.py        # single entry point for all targets
├── new_template.py             # flat workbook rows, dates, cities, GCC relevance score
├── input/
│   ├── template.xlsx           # country list and schema reference
│   └── sources/                # editable source lists (see Sources)
├── feeds/                      # source loading, PDF downloader, template loader
├── extractor/                  # PDF and HTML readers, translator, bucketer, paraphraser
├── exporter/                   # wide workbook and sources report writers
├── tools/                      # scripts that build the source workbooks
├── sources/                    # downloaded PDF cache (large, regenerated by running)
├── output/                     # generated workbooks, paraphrase cache, resume files
├── main.py                     # news digest pipeline (separate, see below)
├── main_global.py, main_usa.py, main_country_states.py,
│   main_country_incentives.py, main_india.py   # older entry points, kept for reference
└── requirements.txt
```

## Troubleshooting

* **Ollama not reachable.** Check `ollama list` shows `llama3.2:3b` and that the Ollama app or service is running, or run with `--nopara`. If paraphrasing is slow, allow parallel requests (on macOS: `launchctl setenv OLLAMA_NUM_PARALLEL 4`, then quit and reopen the Ollama app).
* **A source returned nothing.** See the warning line for the reason. Dead links (404) can be set to `Active` no, blocked sites need another official page.
* **The run stopped with exit code 3.** That is the `--max-minutes` limit. Run the same command again.
* **A PDF gives no text.** It may be image only, so install `ocrmypdf` and Tesseract, or find a text version.
* **The workbook will not save.** Close it in Excel first.

## News digest pipeline (separate)

`main.py` is an earlier, independent pipeline. It reads a Google News RSS search (configured in `config.py`), keeps articles that look GCC relevant, tags state, category, policy and incentive type, stores them in SQLite under `data/`, and exports an Excel report. It shares no code path with `country_extractor.py`.
