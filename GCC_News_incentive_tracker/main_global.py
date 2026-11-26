import argparse
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date

from feeds.policy_downloader import PolicyDownloader
from feeds.previous_workbook_reader import PreviousWorkbookReader
from feeds.sources_loader import SourcesLoader
from feeds.template_loader import TemplateLoader
from extractor.policy_pdf_reader import PolicyPDFReader
from extractor.policy_html_reader import PolicyHTMLReader
from extractor.policy_bucketer import PolicyBucketer
from extractor.paraphraser import OllamaParaphraser
from extractor.translator import ArgosTranslator
from exporter.policy_excel_exporter import PolicyExcelExporter
from exporter.sources_excel_exporter import SourcesExcelExporter


INPUT_DIR = "input"
SOURCES_FILE = os.path.join(INPUT_DIR, "sources.xlsx")
TEMPLATE_FILE = os.path.join(INPUT_DIR, "template.xlsx")

INDIA_SOURCES_DIR = "sources/india"
GLOBAL_SOURCES_DIR = "sources/global"


def empty_buckets():

    return {
        it: {scope: [] for scope in PolicyBucketer.SCOPES}
        for it in PolicyBucketer.INCENTIVE_TYPES
    }


def state_dir_name(state):
    return state.lower().replace(" ", "_")


def country_dir_name(country):
    return country.lower().replace(" ", "_")


def _ensure_one(dest, url, referer, force, label):
    """Return True if the PDF is available at `dest`. Handles both remote
    URL downloads and local-only entries (blank URL → user placed the file)."""

    if not url:
        # Local-only entry: user placed the file manually.
        if os.path.exists(dest):
            print(f"→ {label} (local file, no URL)")
            return True

        print(f"→ {label}: no URL and local file missing at {dest}")
        return False

    print(f"→ {label}")

    return PolicyDownloader.ensure(url, dest, referer=referer, force=force)


DOWNLOAD_WORKERS = 8


def _download_task(jurisdiction, filename, url, referer, dest, force):
    """Return (jurisdiction, dest) if successful, else (jurisdiction, None)."""

    label = f"{jurisdiction} / {filename}"

    ok = _ensure_one(dest, url, referer, force, label)

    return (jurisdiction, dest if ok else None)


def ensure_pdf_sources(india_pdfs, country_pdfs, force=False):

    print(f"Step 1a: ensuring India PDFs ({DOWNLOAD_WORKERS} workers)\n")

    india_available = {state: [] for state in india_pdfs}

    india_tasks = [
        (state, filename, url, referer,
         os.path.join(INDIA_SOURCES_DIR, filename), force)
        for state, pdfs in india_pdfs.items()
        for filename, url, referer in pdfs
    ]

    with ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as ex:

        futures = [ex.submit(_download_task, *t) for t in india_tasks]

        for f in as_completed(futures):

            state, path = f.result()

            if path:
                india_available[state].append(path)

    print(f"\nStep 1b: ensuring country PDFs ({DOWNLOAD_WORKERS} workers)\n")

    country_available = {country: [] for country in country_pdfs}

    country_tasks = [
        (country, filename, url, referer,
         os.path.join(GLOBAL_SOURCES_DIR, country_dir_name(country), filename),
         force)
        for country, pdfs in country_pdfs.items()
        for filename, url, referer in pdfs
    ]

    with ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as ex:

        futures = [ex.submit(_download_task, *t) for t in country_tasks]

        for f in as_completed(futures):

            country, path = f.result()

            if path:
                country_available[country].append(path)

    return india_available, country_available


def build_india_data(india_available, india_html, india_validity, india_states):

    print("\nStep 2a: extracting India\n")

    state_data = {}

    validity_by_state = {}

    for state in india_states:

        validity_by_state[state] = india_validity.get(state, "")

        paths = india_available.get(state, [])

        html_urls = india_html.get(state, [])

        paragraphs = []

        if paths:

            print(f"→ {state}: {len(paths)} PDF(s)")

            for path in paths:

                try:
                    paragraphs.extend(PolicyPDFReader.read(path))
                except Exception as exc:
                    print(f"    ✗ {os.path.basename(path)} failed: {exc}")

        if html_urls:

            print(f"  + {len(html_urls)} HTML reference(s)")

            state_dir = os.path.join(INDIA_SOURCES_DIR, state_dir_name(state))

            for url in html_urls:

                try:
                    paras = PolicyHTMLReader.read(
                        url,
                        country_dir=state_dir,
                        depth=1,
                        pdf_reader=PolicyPDFReader,
                    )
                    paragraphs.extend(paras)
                except Exception:
                    pass

        if not paragraphs:

            state_data[state] = empty_buckets()
            continue

        state_data[state] = PolicyBucketer.bucket(paragraphs)

    return state_data, validity_by_state


EXTRACT_WORKERS = 6


def _extract_one_country(country, paths, html_urls, country_dir):
    """Read PDFs and crawl HTMLs for a single country. Returns
    (country, bucketed_data, paragraph_count).
    Runs safely in a worker thread."""

    paragraphs = []

    for path in paths:

        try:
            paragraphs.extend(PolicyPDFReader.read(path))
        except Exception:
            pass

    for url in html_urls:

        try:
            paras = PolicyHTMLReader.read(
                url, country_dir=country_dir, depth=1,
                pdf_reader=PolicyPDFReader,
            )
            paragraphs.extend(paras)
        except Exception:
            pass

    if not paragraphs:
        return (country, empty_buckets(), 0)

    return (country, PolicyBucketer.bucket(paragraphs), len(paragraphs))


def build_country_data(country_available, country_html, country_validity, all_countries):

    print(f"\nStep 2b: extracting countries ({EXTRACT_WORKERS} workers)\n")

    country_data = {}
    validity_by_country = {}

    tasks = []

    for country in all_countries:

        validity_by_country[country] = country_validity.get(country, "")

        paths = country_available.get(country, [])
        html_urls = country_html.get(country, [])
        country_dir = os.path.join(
            GLOBAL_SOURCES_DIR, country_dir_name(country)
        )

        tasks.append((country, paths, html_urls, country_dir))

    with ThreadPoolExecutor(max_workers=EXTRACT_WORKERS) as ex:

        futures = {
            ex.submit(_extract_one_country, *t): t[0] for t in tasks
        }

        for f in as_completed(futures):

            country, data, para_count = f.result()

            country_data[country] = data

            total = sum(
                len(items)
                for scopes in data.values()
                for items in scopes.values()
            )

            print(f"→ {country}: bucketed {total} items from {para_count} paragraphs")

    return country_data, validity_by_country


def build_sources_rows(srcs, india_available, country_available):

    rows = []

    for state, pdfs in srcs["india_pdfs"].items():

        downloaded = {os.path.basename(p) for p in india_available.get(state, [])}

        for filename, url, _ in pdfs:

            is_local = not url

            status = "Used" if filename in downloaded else "Failed"

            local = (
                os.path.join(INDIA_SOURCES_DIR, filename)
                if filename in downloaded else ""
            )

            if is_local:
                note = ("Local file, extracted" if status == "Used"
                        else "Local file missing at expected path")
            else:
                note = ("Downloaded and extracted" if status == "Used"
                        else "Configured but download failed")

            rows.append((
                "India", state, url or "(local file)", "PDF",
                status, local, note,
            ))

    for state, urls in srcs["india_html"].items():

        for url in urls:

            rows.append((
                "India", state, url, "HTML", "Scraped", "",
                "HTML scraped and bucketed",
            ))

    for country, pdfs in srcs["country_pdfs"].items():

        downloaded = {os.path.basename(p) for p in country_available.get(country, [])}

        for filename, url, _ in pdfs:

            is_local = not url

            status = "Used" if filename in downloaded else "Failed"

            local = (
                os.path.join(GLOBAL_SOURCES_DIR, country_dir_name(country), filename)
                if filename in downloaded else ""
            )

            if is_local:
                note = ("Local file, extracted" if status == "Used"
                        else "Local file missing at expected path")
            else:
                note = ("Downloaded and extracted" if status == "Used"
                        else "Configured but download failed")

            rows.append((
                "Global", country, url or "(local file)", "PDF",
                status, local, note,
            ))

    for country, urls in srcs["country_html"].items():

        for url in urls:

            rows.append((
                "Global", country, url, "HTML", "Scraped", "",
                "HTML scraped and bucketed",
            ))

    return rows


def apply_transform(data, transform_fn, label):
    for jurisdiction, buckets in data.items():

        for incentive_type, scopes in buckets.items():

            for scope, items in scopes.items():

                if not items:
                    continue

                tag = f"{label}/{jurisdiction[:18]}/{incentive_type[:8]}"

                scopes[scope] = transform_fn(items, label=tag)


# Legacy alias for backwards compatibility
apply_paraphrase = apply_transform


def _ordered_keys(*dicts):

    seen = []

    for d in dicts:

        for k in d.keys():

            if k not in seen:
                seen.append(k)

    return seen


def _merge_lists(*lists):

    seen = []

    for lst in lists:

        for item in lst:

            if item not in seen:
                seen.append(item)

    return seen


def output_paths(paraphrased=True):

    today = date.today().isoformat()

    variant = "paraphrased" if paraphrased else "raw"

    return (
        f"output/gcc_incentives_global_{variant}_{today}.xlsx",
        f"output/gcc_sources_{variant}_{today}.xlsx",
    )


def main():

    parser = argparse.ArgumentParser(
        description="Build the global GCC incentives workbook from input/sources.xlsx."
    )

    parser.add_argument(
        "--refresh",
        action="store_true",
        help="Force re-download of all PDFs, ignoring local cache.",
    )

    parser.add_argument(
        "--sources",
        default=SOURCES_FILE,
        help=f"Path to sources workbook (default: {SOURCES_FILE})",
    )

    parser.add_argument(
        "--nopara", "--no-paraphrase",
        dest="paraphrase",
        action="store_false",
        default=True,
        help=(
            "Skip the Ollama paraphrasing step. Cells contain raw extracted "
            "text instead of LLM-cleaned summaries. Useful for quick test "
            "runs or when Ollama is not installed."
        ),
    )

    parser.add_argument(
        "--notranslate", "--no-translate",
        dest="translate",
        action="store_false",
        default=True,
        help=(
            "Skip the argos-translate step. Non-English cells stay in their "
            "original language. Faster runs but cells may not be readable."
        ),
    )

    # Backwards compatible alias; paraphrase is on by default now.
    parser.add_argument(
        "--paraphrase",
        dest="paraphrase",
        action="store_true",
        help=argparse.SUPPRESS,
    )

    parser.add_argument(
        "--model",
        default=OllamaParaphraser.DEFAULT_MODEL,
        help=(
            f"Ollama model name for paraphrasing "
            f"(default: {OllamaParaphraser.DEFAULT_MODEL})"
        ),
    )

    args = parser.parse_args()

    print(f"Loading sources from {args.sources}\n")

    srcs = SourcesLoader.load(args.sources)

    india_states = _ordered_keys(
        srcs["india_pdfs"], srcs["india_html"], srcs["india_validity"]
    )

    countries_from_sources = _ordered_keys(
        srcs["country_pdfs"], srcs["country_html"], srcs["country_validity"]
    )

    countries_from_template = TemplateLoader.load_countries(TEMPLATE_FILE)

    all_countries = _merge_lists(countries_from_template, countries_from_sources)

    print(
        f"  → India states: {len(india_states)}\n"
        f"  → Countries: {len(all_countries)}\n"
    )

    os.makedirs(INDIA_SOURCES_DIR, exist_ok=True)
    os.makedirs(GLOBAL_SOURCES_DIR, exist_ok=True)
    os.makedirs("output", exist_ok=True)

    india_available, country_available = ensure_pdf_sources(
        srcs["india_pdfs"], srcs["country_pdfs"], force=args.refresh,
    )

    state_data, validity_by_state = build_india_data(
        india_available, srcs["india_html"], srcs["india_validity"], india_states,
    )

    country_data, validity_by_country = build_country_data(
        country_available, srcs["country_html"], srcs["country_validity"],
        all_countries,
    )

    if args.translate:

        print("\nStep 2c: translating non-English text via argos-translate\n")

        tr = ArgosTranslator()

        if tr.is_available():

            apply_transform(state_data, tr.translate_many, "states")
            apply_transform(country_data, tr.translate_many, "countries")

    if args.paraphrase:

        print("\nStep 2d: paraphrasing via Ollama\n")

        para = OllamaParaphraser(model=args.model)

        if not para.is_available():

            print(
                "⚠ Ollama not reachable at "
                f"{OllamaParaphraser.HOST} or model "
                f"'{args.model}' not pulled. "
                "Skipping paraphrasing. Run: "
                f"ollama pull {args.model}"
            )

        else:

            apply_transform(state_data, para.paraphrase_many, "states")
            apply_transform(country_data, para.paraphrase_many, "countries")

    print("\nStep 3: writing workbooks\n")

    out_path, sources_path = output_paths(paraphrased=args.paraphrase)

    baseline_prefix = (
        "gcc_incentives_global_paraphrased_"
        if args.paraphrase
        else "gcc_incentives_global_raw_"
    )

    baseline_path = PreviousWorkbookReader.find_latest(
        "output", baseline_prefix, exclude_path=out_path,
    )

    baseline = PreviousWorkbookReader.load(baseline_path)

    if baseline:
        print(f"  ⟂ comparing against baseline: {os.path.basename(baseline_path)}")
    else:
        print("  ⟂ no prior workbook found; all content rendered black")

    PolicyExcelExporter.export_global(
        state_data=state_data,
        validity_by_state=validity_by_state,
        country_data=country_data,
        validity_by_country=validity_by_country,
        placeholder_data={},
        placeholder_validity={},
        output_path=out_path,
        baseline=baseline,
    )

    print(f"  ✓ Wrote {out_path}")

    rows = build_sources_rows(srcs, india_available, country_available)

    SourcesExcelExporter.export(rows, sources_path)

    print(f"  ✓ Wrote {sources_path}")

    # Also emit the flat "new_template" workbook so both formats stay in sync.
    print("\nStep 4: building new_template flat workbook\n")

    try:

        import new_template as _new_template

        _new_template.run()

    except Exception as exc:

        print(f"⚠ new_template.run() failed: {exc}. "
              "You can run 'python3 new_template.py' manually.")


if __name__ == "__main__":
    main()
