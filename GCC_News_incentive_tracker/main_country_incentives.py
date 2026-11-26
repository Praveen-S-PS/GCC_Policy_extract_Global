import argparse
import os
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import date

import openpyxl

from feeds.policy_downloader import PolicyDownloader
from feeds.sources_loader import SourcesLoader
from extractor.policy_pdf_reader import PolicyPDFReader
from extractor.policy_html_reader import PolicyHTMLReader
from extractor.policy_bucketer import PolicyBucketer
from extractor.paraphraser import OllamaParaphraser
from extractor.translator import ArgosTranslator
from exporter.policy_excel_exporter import PolicyExcelExporter

import new_template as nt


INPUT_DIR = "input"
SOURCES_FILE = os.path.join(INPUT_DIR, "sources.xlsx")
GLOBAL_SOURCES_DIR = os.path.join("sources", "global")

DEFAULT_COUNTRIES = ["United Kingdom", "Germany", "France", "Japan"]

DOWNLOAD_WORKERS = 8
EXTRACT_WORKERS = 4
SCAN_WORKERS = int(os.environ.get("SCAN_WORKERS", "10"))


def country_dir_name(country):
    return country.lower().replace(" ", "_")


def empty_buckets():

    return {
        it: {scope: [] for scope in PolicyBucketer.SCOPES}
        for it in PolicyBucketer.INCENTIVE_TYPES
    }


def ensure_pdfs(country, pdfs, force=False):

    available = []

    country_dir = os.path.join(GLOBAL_SOURCES_DIR, country_dir_name(country))

    tasks = [
        (filename, url, referer,
         os.path.join(country_dir, filename))
        for filename, url, referer in pdfs
    ]

    def _one(t):
        filename, url, referer, dest = t
        print(f"  → {country} / {filename}")
        ok = PolicyDownloader.ensure(url, dest, referer=referer, force=force)
        return dest if ok else None

    with ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as ex:
        for dest in ex.map(_one, tasks):
            if dest:
                available.append(dest)

    return available


def extract_country(country, pdf_paths, html_urls):

    country_dir = os.path.join(GLOBAL_SOURCES_DIR, country_dir_name(country))

    paragraphs = []

    for path in pdf_paths:
        try:
            paragraphs.extend(PolicyPDFReader.read(path))
        except Exception as exc:
            print(f"    ✗ {os.path.basename(path)} failed: {exc}")

    for url in html_urls:
        try:
            paragraphs.extend(PolicyHTMLReader.read(
                url, country_dir=country_dir, depth=1,
                pdf_reader=PolicyPDFReader,
            ))
        except Exception as exc:
            print(f"    ✗ {url} failed: {exc}")

    if not paragraphs:
        return empty_buckets(), 0

    return PolicyBucketer.bucket(paragraphs), len(paragraphs)


def apply_transform(buckets, transform_fn, label):

    for incentive_type, scopes in buckets.items():

        for scope, items in scopes.items():

            if not items:
                continue

            tag = f"{label}/{incentive_type[:12]}"

            scopes[scope] = transform_fn(items, label=tag)


def scan_country_dates(country, pdf_paths, html_urls):

    tasks = [(p, False) for p in pdf_paths] + [(u, True) for u in html_urls[:3]]

    if not tasks:
        return ("", "", "", "")

    lock = threading.Lock()
    best = [("", "", "", "")]

    def _run(t):
        target, is_url = t
        dates = nt.scan_source_for_validity(target, is_url=is_url)
        if dates[0] or dates[2]:
            with lock:
                if not (best[0][0] or best[0][2]):
                    best[0] = dates

    with ThreadPoolExecutor(max_workers=min(SCAN_WORKERS, max(1, len(tasks)))) as ex:
        list(ex.map(_run, tasks))

    return best[0]


def build_country_rows(country, buckets, source_dates):

    rows = []

    for incentive_type, scopes in buckets.items():

        section = nt.SECTION_MAP.get(incentive_type)

        if not section:
            continue

        for scope_label, items in scopes.items():

            if not items:
                continue

            scope_type, tier_scope = nt.SCOPE_MAP_GLOBAL.get(
                scope_label, ("Country", scope_label)
            )

            for b in items:

                text_dates = nt.extract_validity_from_text(b)

                start, s_prec, end, e_prec = nt._apply_source_fallback(
                    text_dates, source_dates, ("", "", "", ""),
                )

                if nt.is_expired(end, e_prec):
                    continue

                cities = nt.detect_cities(b, nt.COUNTRY_CITIES.get(country, {}))

                relevance = nt.compute_gcc_relevance(b, country)

                if cities:
                    for city_name, tier in cities:
                        rows.append({
                            "section_name": section,
                            "scope_type": "City",
                            "country_name": country,
                            "state_name": "",
                            "city_name": city_name,
                            "tier_scope": tier,
                            "description": b,
                            "validity_start_date": start,
                            "validity_start_precision": s_prec,
                            "validity_end_date": end,
                            "validity_end_precision": e_prec,
                            "gcc_relevance": relevance,
                        })
                else:
                    rows.append({
                        "section_name": section,
                        "scope_type": scope_type,
                        "country_name": country,
                        "state_name": "",
                        "city_name": "",
                        "tier_scope": tier_scope,
                        "description": b,
                        "validity_start_date": start,
                        "validity_start_precision": s_prec,
                        "validity_end_date": end,
                        "validity_end_precision": e_prec,
                        "gcc_relevance": relevance,
                    })

    rows.sort(key=lambda r: -(r.get("gcc_relevance") or 0))

    return rows


def write_country_files(country, buckets, validity, rows, output_dir):

    os.makedirs(output_dir, exist_ok=True)

    today = date.today().isoformat()

    # 1. Wide-format single-country workbook (audit trail / matches the
    # style of the shared Global sheet, reusing PolicyExcelExporter as-is).
    wide_path = os.path.join(output_dir, f"gcc_incentives_{country_dir_name(country)}_{today}.xlsx")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    PolicyExcelExporter._write_definition_sheet(wb)
    PolicyExcelExporter._write_country_sheet(wb, country, buckets, validity)
    wb.save(wide_path)

    # 2. Flat new_template-format workbook, split by relevance score.
    flat_path = os.path.join(output_dir, f"new_template_{country_dir_name(country)}_policy_{today}.xlsx")

    signal_rows = [r for r in rows if int(r.get("gcc_relevance") or 0) >= 1]
    other_rows = [r for r in rows if int(r.get("gcc_relevance") or 0) == 0]

    wb2 = openpyxl.Workbook()
    wb2.remove(wb2.active)
    nt._write_flat_sheet(wb2, country[:31], signal_rows)
    nt._write_flat_sheet(wb2, f"Other - {country}"[:31], other_rows)
    wb2.save(flat_path)

    return wide_path, flat_path, len(signal_rows), len(other_rows)


def main():

    parser = argparse.ArgumentParser(
        description="Build a per-country GCC incentives workbook for a small list of countries."
    )

    parser.add_argument(
        "--countries", nargs="+", default=DEFAULT_COUNTRIES,
        help=f"Countries to process (default: {DEFAULT_COUNTRIES})",
    )

    parser.add_argument("--refresh", action="store_true")

    parser.add_argument(
        "--nopara", "--no-paraphrase", dest="paraphrase",
        action="store_false", default=True,
    )

    parser.add_argument(
        "--notranslate", "--no-translate", dest="translate",
        action="store_false", default=True,
    )

    parser.add_argument("--model", default=OllamaParaphraser.DEFAULT_MODEL)

    parser.add_argument(
        "--min-len", type=int, default=0,
        help=(
            "Minimum bullet length to send to Ollama (default: 0, meaning "
            "every bullet is paraphrased). The main pipelines default this "
            "higher (250) to skip short bullets for speed; this script "
            "defaults to paraphrasing everything since these are small, "
            "one-off per-country runs."
        ),
    )

    args = parser.parse_args()

    srcs = SourcesLoader.load(SOURCES_FILE)

    para = (
        OllamaParaphraser(model=args.model, min_len=args.min_len)
        if args.paraphrase else None
    )
    para_ok = para.is_available() if para else False

    if args.paraphrase and not para_ok:
        print(f"⚠ Ollama not reachable/model not pulled; continuing without paraphrasing.")

    tr = ArgosTranslator() if args.translate else None
    tr_ok = tr.is_available() if tr else False

    for country in args.countries:

        print(f"\n=== {country} ===")

        pdfs = srcs["country_pdfs"].get(country, [])
        html_urls = srcs["country_html"].get(country, [])
        validity = srcs["country_validity"].get(country, "")

        if not pdfs and not html_urls:
            print(f"  ⚠ no sources configured for {country} in {SOURCES_FILE}; skipping.")
            continue

        pdf_paths = ensure_pdfs(country, pdfs, force=args.refresh) if pdfs else []

        buckets, para_count = extract_country(country, pdf_paths, html_urls)

        total = sum(len(items) for scopes in buckets.values() for items in scopes.values())
        print(f"  bucketed {total} items from {para_count} paragraphs")

        if tr_ok:
            apply_transform(buckets, tr.translate_many, country[:14])

        if para_ok:
            apply_transform(buckets, para.paraphrase_many, country[:14])

        source_dates = scan_country_dates(country, pdf_paths, html_urls)

        rows = build_country_rows(country, buckets, source_dates)

        output_dir = os.path.join("output", country)

        wide_path, flat_path, signal_n, other_n = write_country_files(
            country, buckets, validity, rows, output_dir,
        )

        print(f"  ✓ Wrote {wide_path}")
        print(f"  ✓ Wrote {flat_path}  ({country}: {signal_n} rows, Other - {country}: {other_n} rows)")


if __name__ == "__main__":
    main()
