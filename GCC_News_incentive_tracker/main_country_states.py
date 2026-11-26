"""Builds a state/region/prefecture-level GCC incentives workbook for a
single country, the same treatment main_usa.py gives the 50 US states.

Each supported country has its own sources file (input/sources_<slug>.xlsx,
same 8-column schema as sources.xlsx/sources_usa.xlsx) and its own city
registry for Tier 1/2/3 detection. Output goes to output/<Country>/ as a
wide-format audit workbook plus a flat new_template-format workbook split
by GCC relevance score.

None of these countries have a dedicated "Global Capability Centre" policy
(confirmed via separate research), so -- like most of the existing Global
sheet and the USA build -- this treats general state/region economic
development and investment-promotion content as proxy data.

Standalone: does not import or modify main_global.py, main_usa.py, or
new_template.py's own pipeline, only reuses their pure helper functions.
"""

import argparse
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font

from feeds.policy_downloader import PolicyDownloader
from feeds.previous_workbook_reader import PreviousWorkbookReader
from feeds.sources_loader import SourcesLoader
from extractor.policy_pdf_reader import PolicyPDFReader
from extractor.policy_html_reader import PolicyHTMLReader
from extractor.policy_bucketer import PolicyBucketer
from extractor.paraphraser import OllamaParaphraser
from extractor.translator import ArgosTranslator
from exporter.policy_excel_exporter import PolicyExcelExporter

import new_template as nt


INPUT_DIR = "input"
SOURCES_DIR_PREFIX = os.path.join("sources")
DOWNLOAD_WORKERS = 8
EXTRACT_WORKERS = 6
SCAN_WORKERS = int(os.environ.get("SCAN_WORKERS", "10"))


def slug(name):
    return name.lower().replace(" ", "_").replace("-", "_").replace("'", "")


# ---------------------------------------------------------------------------
# Per-country configuration: the official list of states/regions/prefectures
# and a hand-curated Tier 1 / Tier 2 & 3 city registry for each, at the same
# level of detail as new_template.py's existing COUNTRY_CITIES entries.
# ---------------------------------------------------------------------------

COUNTRY_STATES = {
    "United Kingdom": ["England", "Scotland", "Wales", "Northern Ireland"],
    "Germany": [
        "Baden-Württemberg", "Bavaria", "Berlin", "Brandenburg", "Bremen",
        "Hamburg", "Hesse", "Lower Saxony", "Mecklenburg-Vorpommern",
        "North Rhine-Westphalia", "Rhineland-Palatinate", "Saarland",
        "Saxony", "Saxony-Anhalt", "Schleswig-Holstein", "Thuringia",
    ],
    "France": [
        "Île-de-France", "Auvergne-Rhône-Alpes", "Hauts-de-France",
        "Grand Est", "Provence-Alpes-Côte d'Azur", "Occitanie",
        "Nouvelle-Aquitaine", "Bretagne", "Normandie",
        "Bourgogne-Franche-Comté", "Centre-Val de Loire", "Pays de la Loire",
        "Corse",
    ],
    "Japan": [
        "Hokkaido", "Aomori", "Iwate", "Miyagi", "Akita", "Yamagata",
        "Fukushima", "Ibaraki", "Tochigi", "Gunma", "Saitama", "Chiba",
        "Tokyo", "Kanagawa", "Niigata", "Toyama", "Ishikawa", "Fukui",
        "Yamanashi", "Nagano", "Gifu", "Shizuoka", "Aichi", "Mie", "Shiga",
        "Kyoto", "Osaka", "Hyogo", "Nara", "Wakayama", "Tottori", "Shimane",
        "Okayama", "Hiroshima", "Yamaguchi", "Tokushima", "Kagawa", "Ehime",
        "Kochi", "Fukuoka", "Saga", "Nagasaki", "Kumamoto", "Oita",
        "Miyazaki", "Kagoshima", "Okinawa",
    ],
}

STATE_CITIES = {
    "United Kingdom": {
        "England": {"Tier 1": ["London"], "Tier 2 & 3": ["Manchester", "Birmingham", "Leeds", "Bristol"]},
        "Scotland": {"Tier 1": ["Edinburgh", "Glasgow"], "Tier 2 & 3": ["Aberdeen", "Dundee"]},
        "Wales": {"Tier 1": ["Cardiff"], "Tier 2 & 3": ["Swansea", "Newport"]},
        "Northern Ireland": {"Tier 1": ["Belfast"], "Tier 2 & 3": ["Derry", "Lisburn"]},
    },
    "Germany": {
        "Baden-Württemberg": {"Tier 1": ["Stuttgart"], "Tier 2 & 3": ["Karlsruhe", "Mannheim", "Freiburg"]},
        "Bavaria": {"Tier 1": ["Munich", "München"], "Tier 2 & 3": ["Nuremberg", "Nürnberg", "Augsburg"]},
        "Berlin": {"Tier 1": ["Berlin"], "Tier 2 & 3": []},
        "Brandenburg": {"Tier 1": ["Potsdam"], "Tier 2 & 3": ["Cottbus"]},
        "Bremen": {"Tier 1": ["Bremen"], "Tier 2 & 3": []},
        "Hamburg": {"Tier 1": ["Hamburg"], "Tier 2 & 3": []},
        "Hesse": {"Tier 1": ["Frankfurt"], "Tier 2 & 3": ["Wiesbaden", "Darmstadt"]},
        "Lower Saxony": {"Tier 1": ["Hanover", "Hannover"], "Tier 2 & 3": ["Braunschweig", "Osnabrück"]},
        "Mecklenburg-Vorpommern": {"Tier 1": ["Rostock"], "Tier 2 & 3": ["Schwerin"]},
        "North Rhine-Westphalia": {"Tier 1": ["Cologne", "Köln", "Düsseldorf"], "Tier 2 & 3": ["Dortmund", "Essen", "Bonn"]},
        "Rhineland-Palatinate": {"Tier 1": ["Mainz"], "Tier 2 & 3": ["Ludwigshafen", "Koblenz"]},
        "Saarland": {"Tier 1": ["Saarbrücken"], "Tier 2 & 3": []},
        "Saxony": {"Tier 1": ["Dresden", "Leipzig"], "Tier 2 & 3": ["Chemnitz"]},
        "Saxony-Anhalt": {"Tier 1": ["Magdeburg"], "Tier 2 & 3": ["Halle"]},
        "Schleswig-Holstein": {"Tier 1": ["Kiel"], "Tier 2 & 3": ["Lübeck"]},
        "Thuringia": {"Tier 1": ["Erfurt"], "Tier 2 & 3": ["Jena"]},
    },
    "France": {
        "Île-de-France": {"Tier 1": ["Paris"], "Tier 2 & 3": []},
        "Auvergne-Rhône-Alpes": {"Tier 1": ["Lyon"], "Tier 2 & 3": ["Grenoble", "Saint-Étienne", "Clermont-Ferrand"]},
        "Hauts-de-France": {"Tier 1": ["Lille"], "Tier 2 & 3": ["Amiens"]},
        "Grand Est": {"Tier 1": ["Strasbourg"], "Tier 2 & 3": ["Reims", "Metz", "Nancy"]},
        "Provence-Alpes-Côte d'Azur": {"Tier 1": ["Marseille", "Nice"], "Tier 2 & 3": ["Toulon", "Aix-en-Provence"]},
        "Occitanie": {"Tier 1": ["Toulouse"], "Tier 2 & 3": ["Montpellier"]},
        "Nouvelle-Aquitaine": {"Tier 1": ["Bordeaux"], "Tier 2 & 3": ["Poitiers", "Limoges"]},
        "Bretagne": {"Tier 1": ["Rennes"], "Tier 2 & 3": ["Brest"]},
        "Normandie": {"Tier 1": ["Rouen"], "Tier 2 & 3": ["Caen", "Le Havre"]},
        "Bourgogne-Franche-Comté": {"Tier 1": ["Dijon"], "Tier 2 & 3": ["Besançon"]},
        "Centre-Val de Loire": {"Tier 1": ["Orléans"], "Tier 2 & 3": ["Tours"]},
        "Pays de la Loire": {"Tier 1": ["Nantes"], "Tier 2 & 3": ["Angers", "Le Mans"]},
        "Corse": {"Tier 1": ["Ajaccio"], "Tier 2 & 3": ["Bastia"]},
    },
    "Japan": {
        "Hokkaido": {"Tier 1": ["Sapporo"], "Tier 2 & 3": []},
        "Aomori": {"Tier 1": ["Aomori"], "Tier 2 & 3": []},
        "Iwate": {"Tier 1": ["Morioka"], "Tier 2 & 3": []},
        "Miyagi": {"Tier 1": ["Sendai"], "Tier 2 & 3": []},
        "Akita": {"Tier 1": ["Akita"], "Tier 2 & 3": []},
        "Yamagata": {"Tier 1": ["Yamagata"], "Tier 2 & 3": []},
        "Fukushima": {"Tier 1": ["Fukushima"], "Tier 2 & 3": []},
        "Ibaraki": {"Tier 1": ["Mito"], "Tier 2 & 3": []},
        "Tochigi": {"Tier 1": ["Utsunomiya"], "Tier 2 & 3": []},
        "Gunma": {"Tier 1": ["Maebashi"], "Tier 2 & 3": []},
        "Saitama": {"Tier 1": ["Saitama"], "Tier 2 & 3": []},
        "Chiba": {"Tier 1": ["Chiba"], "Tier 2 & 3": []},
        "Tokyo": {"Tier 1": ["Tokyo"], "Tier 2 & 3": []},
        "Kanagawa": {"Tier 1": ["Yokohama"], "Tier 2 & 3": ["Kawasaki"]},
        "Niigata": {"Tier 1": ["Niigata"], "Tier 2 & 3": []},
        "Toyama": {"Tier 1": ["Toyama"], "Tier 2 & 3": []},
        "Ishikawa": {"Tier 1": ["Kanazawa"], "Tier 2 & 3": []},
        "Fukui": {"Tier 1": ["Fukui"], "Tier 2 & 3": []},
        "Yamanashi": {"Tier 1": ["Kofu"], "Tier 2 & 3": []},
        "Nagano": {"Tier 1": ["Nagano"], "Tier 2 & 3": ["Matsumoto"]},
        "Gifu": {"Tier 1": ["Gifu"], "Tier 2 & 3": []},
        "Shizuoka": {"Tier 1": ["Shizuoka"], "Tier 2 & 3": ["Hamamatsu"]},
        "Aichi": {"Tier 1": ["Nagoya"], "Tier 2 & 3": []},
        "Mie": {"Tier 1": ["Tsu"], "Tier 2 & 3": []},
        "Shiga": {"Tier 1": ["Otsu"], "Tier 2 & 3": []},
        "Kyoto": {"Tier 1": ["Kyoto"], "Tier 2 & 3": []},
        "Osaka": {"Tier 1": ["Osaka"], "Tier 2 & 3": []},
        "Hyogo": {"Tier 1": ["Kobe"], "Tier 2 & 3": []},
        "Nara": {"Tier 1": ["Nara"], "Tier 2 & 3": []},
        "Wakayama": {"Tier 1": ["Wakayama"], "Tier 2 & 3": []},
        "Tottori": {"Tier 1": ["Tottori"], "Tier 2 & 3": []},
        "Shimane": {"Tier 1": ["Matsue"], "Tier 2 & 3": []},
        "Okayama": {"Tier 1": ["Okayama"], "Tier 2 & 3": []},
        "Hiroshima": {"Tier 1": ["Hiroshima"], "Tier 2 & 3": []},
        "Yamaguchi": {"Tier 1": ["Yamaguchi"], "Tier 2 & 3": []},
        "Tokushima": {"Tier 1": ["Tokushima"], "Tier 2 & 3": []},
        "Kagawa": {"Tier 1": ["Takamatsu"], "Tier 2 & 3": []},
        "Ehime": {"Tier 1": ["Matsuyama"], "Tier 2 & 3": []},
        "Kochi": {"Tier 1": ["Kochi"], "Tier 2 & 3": []},
        "Fukuoka": {"Tier 1": ["Fukuoka"], "Tier 2 & 3": ["Kitakyushu"]},
        "Saga": {"Tier 1": ["Saga"], "Tier 2 & 3": []},
        "Nagasaki": {"Tier 1": ["Nagasaki"], "Tier 2 & 3": []},
        "Kumamoto": {"Tier 1": ["Kumamoto"], "Tier 2 & 3": []},
        "Oita": {"Tier 1": ["Oita"], "Tier 2 & 3": []},
        "Miyazaki": {"Tier 1": ["Miyazaki"], "Tier 2 & 3": []},
        "Kagoshima": {"Tier 1": ["Kagoshima"], "Tier 2 & 3": []},
        "Okinawa": {"Tier 1": ["Naha"], "Tier 2 & 3": []},
    },
}


def sources_file_for(country):
    return os.path.join(INPUT_DIR, f"sources_{slug(country)}.xlsx")


def sources_dir_for(country):
    return os.path.join(SOURCES_DIR_PREFIX, slug(country))


def output_dir_for(country):
    return os.path.join("output", country)


def empty_buckets():

    return {
        it: {scope: [] for scope in PolicyBucketer.SCOPES}
        for it in PolicyBucketer.INCENTIVE_TYPES
    }


def _download_task(state, filename, url, referer, dest, force):

    label = f"{state} / {filename}"

    if not url:
        ok = os.path.exists(dest)
        print(f"→ {label} (local file, no URL)" if ok else f"→ {label}: no URL and file missing")
    else:
        print(f"→ {label}")
        ok = PolicyDownloader.ensure(url, dest, referer=referer, force=force)

    return (state, dest if ok else None)


def ensure_pdf_sources(country, state_pdfs, force=False):

    available = {state: [] for state in state_pdfs}

    tasks = [
        (state, filename, url, referer,
         os.path.join(sources_dir_for(country), slug(state), filename), force)
        for state, pdfs in state_pdfs.items()
        for filename, url, referer in pdfs
    ]

    if not tasks:
        return available

    with ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as ex:

        futures = [ex.submit(_download_task, *t) for t in tasks]

        for f in as_completed(futures):
            state, path = f.result()
            if path:
                available[state].append(path)

    return available


def _extract_one_state(country, state, paths, html_urls):

    state_dir = os.path.join(sources_dir_for(country), slug(state))

    paragraphs = []

    for path in paths:
        try:
            paragraphs.extend(PolicyPDFReader.read(path))
        except Exception:
            pass

    for url in html_urls:
        try:
            paragraphs.extend(PolicyHTMLReader.read(
                url, country_dir=state_dir, depth=1, pdf_reader=PolicyPDFReader,
            ))
        except Exception:
            pass

    return (state, paragraphs)


def build_state_data(country, state_available, state_html, state_validity, states,
                      translator=None):
    """Extract raw paragraphs per state, translate to English, THEN bucket.

    Bucketing must happen after translation, not before: PolicyBucketer's
    keyword vocabulary is English/Spanish only, so feeding it raw German/
    French/Japanese text drops almost everything at the classification
    step. Confirmed directly on a real source (Hesse's WIBank page: 314
    extracted paragraphs, 0 bucketed items under the old bucket-then-
    translate order; 16 bucketed items once translated first)."""

    validity_by_state = {state: state_validity.get(state, "") for state in states}

    tasks = [
        (country, state, state_available.get(state, []), state_html.get(state, []))
        for state in states
    ]

    raw_by_state = {}

    with ThreadPoolExecutor(max_workers=EXTRACT_WORKERS) as ex:

        futures = {ex.submit(_extract_one_state, *t): t[1] for t in tasks}

        for f in as_completed(futures):

            state, paragraphs = f.result()
            raw_by_state[state] = paragraphs

    state_data = {}

    for state, paragraphs in raw_by_state.items():

        if not paragraphs:
            state_data[state] = empty_buckets()
            print(f"→ {state}: bucketed 0 items from 0 paragraphs")
            continue

        texts = [p["text"] for p in paragraphs]

        if translator is not None:
            texts = translator.translate_many(texts, label=f"{state[:14]}/raw")

        translated_paragraphs = [
            {"heading": p.get("heading", ""), "text": t}
            for p, t in zip(paragraphs, texts)
        ]

        data = PolicyBucketer.bucket(translated_paragraphs)

        state_data[state] = data

        total = sum(len(items) for scopes in data.values() for items in scopes.values())

        print(f"→ {state}: bucketed {total} items from {len(paragraphs)} paragraphs")

    return state_data, validity_by_state


def apply_transform(data, transform_fn, label):

    for jurisdiction, buckets in data.items():

        for incentive_type, scopes in buckets.items():

            for scope, items in scopes.items():

                if not items:
                    continue

                tag = f"{label}/{jurisdiction[:14]}/{incentive_type[:8]}"

                scopes[scope] = transform_fn(items, label=tag)


def build_source_dates_index(country, srcs):

    dates_by_state = {}
    lock = threading.Lock()

    def _remember(state, dates):
        with lock:
            if not dates_by_state.get(state) and (dates[0] or dates[2]):
                dates_by_state[state] = dates

    tasks = []

    for state, pdfs in srcs["country_pdfs"].items():
        for filename, _, _ in pdfs:
            tasks.append((state, os.path.join(
                sources_dir_for(country), slug(state), filename), False))

    for state, urls in srcs["country_html"].items():
        for url in urls[:2]:
            if url:
                tasks.append((state, url, True))

    if not tasks:
        return dates_by_state

    print(f"  ⓘ scanning source files for validity dates ({SCAN_WORKERS} workers)")

    def _run(task):
        state, target, is_url = task
        dates = nt.scan_source_for_validity(target, is_url=is_url)
        _remember(state, dates)

    with ThreadPoolExecutor(max_workers=SCAN_WORKERS) as ex:
        list(ex.map(_run, tasks))

    return dates_by_state


def extract_state_rows(country, state_data, validity_by_state, source_dates_index):

    rows_by_state = {state: [] for state in state_data}
    city_registry = STATE_CITIES.get(country, {})

    for state, buckets in state_data.items():

        validity_start, vs_prec, validity_end, ve_prec = nt.parse_validity(
            validity_by_state.get(state, "")
        )

        for incentive_type, scopes in buckets.items():

            section = nt.SECTION_MAP.get(incentive_type)

            if not section:
                continue

            for scope_label, items in scopes.items():

                if not items:
                    continue

                scope_type, tier_scope = nt.SCOPE_MAP_INDIA.get(
                    scope_label, ("State", scope_label)
                )

                for b in items:

                    text_dates = nt.extract_validity_from_text(b)

                    start, s_prec, end, e_prec = nt._apply_source_fallback(
                        text_dates,
                        source_dates_index.get(state),
                        (validity_start, vs_prec, validity_end, ve_prec),
                    )

                    if nt.is_expired(end, e_prec):
                        continue

                    cities = nt.detect_cities(b, city_registry.get(state, {}))

                    relevance = nt.compute_gcc_relevance(b, country)

                    if cities:
                        for city_name, tier in cities:
                            rows_by_state[state].append({
                                "section_name": section,
                                "scope_type": "City",
                                "country_name": country,
                                "state_name": state,
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
                        rows_by_state[state].append({
                            "section_name": section,
                            "scope_type": scope_type,
                            "country_name": country,
                            "state_name": state,
                            "city_name": "",
                            "tier_scope": tier_scope,
                            "description": b,
                            "validity_start_date": start,
                            "validity_start_precision": s_prec,
                            "validity_end_date": end,
                            "validity_end_precision": e_prec,
                            "gcc_relevance": relevance,
                        })

    return rows_by_state


# ---------------------------------------------------------------------------
# Excel export: reuses PolicyExcelExporter's diff/style building blocks.
# ---------------------------------------------------------------------------

_PEE = PolicyExcelExporter


def _write_wide_sheet(wb, country, state_data, validity_by_state):

    ws = wb.create_sheet(f"{country}-State wise"[:31])

    states = list(state_data.keys())

    headers = ["Incentive Type", "Applicable in which part of the state/region"] + states

    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = _PEE.HEADER_FONT
        cell.fill = _PEE.HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = _PEE.BORDER

    ws.cell(row=2, column=1, value="Policy Validity Period ->").font = Font(bold=True)
    ws.cell(row=2, column=1).alignment = _PEE.WRAP
    ws.cell(row=2, column=1).border = _PEE.BORDER
    ws.cell(row=2, column=2, value="").border = _PEE.BORDER

    for col_idx, state in enumerate(states, start=3):
        cell = ws.cell(row=2, column=col_idx, value=validity_by_state.get(state, ""))
        cell.alignment = _PEE.WRAP
        cell.border = _PEE.BORDER

    row_cursor = 3

    for incentive_type in _PEE.INCENTIVE_TYPES:

        type_start_row = row_cursor

        for scope in _PEE.SCOPES:

            type_cell = ws.cell(row=row_cursor, column=1)

            if scope == _PEE.SCOPES[0]:
                type_cell.value = incentive_type
                type_cell.font = Font(bold=True)
                type_cell.fill = _PEE.TYPE_FILL

            type_cell.alignment = _PEE.WRAP
            type_cell.border = _PEE.BORDER

            scope_cell = ws.cell(row=row_cursor, column=2, value=scope)
            scope_cell.alignment = _PEE.WRAP
            scope_cell.border = _PEE.BORDER

            for col_idx, state in enumerate(states, start=3):

                items = state_data.get(state, {}).get(incentive_type, {}).get(scope, [])

                _PEE._write_diff_cell(
                    ws, row_cursor, col_idx, items, "",
                    has_baseline=False, new_items=None,
                )

            row_cursor += 1

        if type_start_row + 1 <= row_cursor - 1:
            ws.merge_cells(start_row=type_start_row, start_column=1,
                            end_row=row_cursor - 1, end_column=1)

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 28

    for col_idx in range(3, 3 + len(states)):
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[col_letter].width = 45

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30

    for r in range(3, row_cursor):
        ws.row_dimensions[r].height = 200


def write_outputs(country, state_data, validity_by_state, rows_by_state):

    output_dir = output_dir_for(country)
    os.makedirs(output_dir, exist_ok=True)

    today = date.today().isoformat()

    wide_path = os.path.join(output_dir, f"gcc_incentives_{slug(country)}_states_{today}.xlsx")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _PEE._write_definition_sheet(wb)
    _write_wide_sheet(wb, country, state_data, validity_by_state)
    wb.save(wide_path)

    all_rows = []
    for state in rows_by_state:
        all_rows.extend(sorted(rows_by_state[state], key=lambda r: -(r.get("gcc_relevance") or 0)))

    signal_rows = [r for r in all_rows if int(r.get("gcc_relevance") or 0) >= 1]
    other_rows = [r for r in all_rows if int(r.get("gcc_relevance") or 0) == 0]

    flat_path = os.path.join(output_dir, f"new_template_{slug(country)}_states_policy_{today}.xlsx")

    wb2 = openpyxl.Workbook()
    wb2.remove(wb2.active)
    nt._write_flat_sheet(wb2, country[:31], signal_rows)
    nt._write_flat_sheet(wb2, f"Other - {country}"[:31], other_rows)
    wb2.save(flat_path)

    return wide_path, flat_path, len(signal_rows), len(other_rows)


def main():

    parser = argparse.ArgumentParser(
        description="Build a state/region/prefecture-level GCC incentives workbook for one country."
    )

    parser.add_argument("--country", required=True, choices=list(COUNTRY_STATES.keys()))
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument("--nopara", "--no-paraphrase", dest="paraphrase", action="store_false", default=True)
    parser.add_argument("--notranslate", "--no-translate", dest="translate", action="store_false", default=True)
    parser.add_argument("--model", default=OllamaParaphraser.DEFAULT_MODEL)
    parser.add_argument("--min-len", type=int, default=0)

    args = parser.parse_args()

    country = args.country
    sources_file = sources_file_for(country)

    if not os.path.exists(sources_file):
        print(f"✗ {sources_file} not found. Build it first (see tools/build_sources_<country>_workbook.py).")
        return

    print(f"Loading {country} sources from {sources_file}\n")

    srcs = SourcesLoader.load(sources_file)

    states = list(srcs["country_pdfs"].keys())
    for s in srcs["country_html"]:
        if s not in states:
            states.append(s)
    for s in COUNTRY_STATES[country]:
        if s not in states:
            states.append(s)

    os.makedirs(sources_dir_for(country), exist_ok=True)

    state_available = ensure_pdf_sources(country, srcs["country_pdfs"], force=args.refresh)

    translator = None

    if args.translate:
        tr = ArgosTranslator()
        if tr.is_available():
            translator = tr
        else:
            print("⚠ argostranslate not available; non-English content will not be translated.")

    print(f"\nExtracting {country} states/regions ({EXTRACT_WORKERS} workers)\n")

    state_data, validity_by_state = build_state_data(
        country, state_available, srcs["country_html"], srcs["country_validity"], states,
        translator=translator,
    )

    if args.paraphrase:
        para = OllamaParaphraser(model=args.model, min_len=args.min_len)
        if para.is_available():
            print("\nParaphrasing via Ollama\n")
            apply_transform(state_data, para.paraphrase_many, country[:10])
        else:
            print(f"⚠ Ollama not reachable/model not pulled; skipping paraphrasing.")

    source_dates_index = build_source_dates_index(country, srcs)

    rows_by_state = extract_state_rows(country, state_data, validity_by_state, source_dates_index)

    wide_path, flat_path, signal_n, other_n = write_outputs(
        country, state_data, validity_by_state, rows_by_state,
    )

    print(f"\n✓ Wrote {wide_path}")
    print(f"✓ Wrote {flat_path}  ({country}: {signal_n} rows, Other - {country}: {other_n} rows)")


if __name__ == "__main__":
    main()
