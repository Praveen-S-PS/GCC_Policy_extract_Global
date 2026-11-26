"""Builds the United States GCC incentives workbook, state by state.

Standalone counterpart to main_global.py / new_template.py, scoped to the
50 US states. Does not import or modify either of those files, and writes
its own output under output/United States/ so it cannot collide with the
India/Global run.

US states do not have policies that name Global Capability Centres the way
India, Mexico, and Poland do, so this treats each state's general business
investment incentive programs (tax credits, grants, site selection
incentives) as proxy content, the same approach already used for most
countries in the Global sheet.
"""

import argparse
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from feeds.policy_downloader import PolicyDownloader
from feeds.previous_workbook_reader import PreviousWorkbookReader
from feeds.sources_loader import SourcesLoader
from extractor.policy_pdf_reader import PolicyPDFReader
from extractor.policy_html_reader import PolicyHTMLReader
from extractor.policy_bucketer import PolicyBucketer
from extractor.paraphraser import OllamaParaphraser
from extractor.translator import ArgosTranslator
from exporter.policy_excel_exporter import PolicyExcelExporter
from exporter.sources_excel_exporter import SourcesExcelExporter

# Reused as a library of pure helper functions (date extraction, city
# detection, relevance scoring, flat-sheet writer). Importing this module
# does not execute or alter its own pipeline in any way.
import new_template as nt


INPUT_DIR = "input"
SOURCES_FILE = os.path.join(INPUT_DIR, "sources_usa.xlsx")

USA_SOURCES_DIR = os.path.join("sources", "usa")
OUTPUT_DIR = os.path.join("output", "United States")

DOWNLOAD_WORKERS = 8
EXTRACT_WORKERS = 6
SCAN_WORKERS = int(os.environ.get("SCAN_WORKERS", "10"))


US_STATES = [
    "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
    "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
    "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana",
    "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
    "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada",
    "New Hampshire", "New Jersey", "New Mexico", "New York",
    "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon",
    "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota",
    "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington",
    "West Virginia", "Wisconsin", "Wyoming",
]


# Major business-hub cities per state, used to upgrade scope_type to "City"
# in the flat output. Hand-curated from well-known metro areas, the same
# level of detail as the existing COUNTRY_CITIES registry in new_template.py.
US_STATE_CITIES = {
    "Alabama": {"Tier 1": ["Birmingham"], "Tier 2 & 3": ["Huntsville", "Montgomery", "Mobile"]},
    "Alaska": {"Tier 1": ["Anchorage"], "Tier 2 & 3": ["Fairbanks", "Juneau"]},
    "Arizona": {"Tier 1": ["Phoenix", "Scottsdale", "Tempe"], "Tier 2 & 3": ["Tucson", "Mesa", "Chandler"]},
    "Arkansas": {"Tier 1": ["Little Rock"], "Tier 2 & 3": ["Fayetteville", "Springdale", "Bentonville"]},
    "California": {"Tier 1": ["San Francisco", "Los Angeles", "San Jose", "Silicon Valley"], "Tier 2 & 3": ["San Diego", "Sacramento", "Oakland", "Irvine"]},
    "Colorado": {"Tier 1": ["Denver"], "Tier 2 & 3": ["Boulder", "Colorado Springs", "Fort Collins"]},
    "Connecticut": {"Tier 1": ["Stamford", "Hartford"], "Tier 2 & 3": ["New Haven", "Bridgeport"]},
    "Delaware": {"Tier 1": ["Wilmington"], "Tier 2 & 3": ["Dover", "Newark"]},
    "Florida": {"Tier 1": ["Miami", "Tampa", "Orlando"], "Tier 2 & 3": ["Jacksonville", "Fort Lauderdale", "St. Petersburg"]},
    "Georgia": {"Tier 1": ["Atlanta"], "Tier 2 & 3": ["Savannah", "Augusta", "Columbus"]},
    "Hawaii": {"Tier 1": ["Honolulu"], "Tier 2 & 3": ["Hilo", "Kailua"]},
    "Idaho": {"Tier 1": ["Boise"], "Tier 2 & 3": ["Idaho Falls", "Coeur d'Alene"]},
    "Illinois": {"Tier 1": ["Chicago"], "Tier 2 & 3": ["Naperville", "Aurora", "Rockford"]},
    "Indiana": {"Tier 1": ["Indianapolis"], "Tier 2 & 3": ["Fort Wayne", "Evansville", "South Bend"]},
    "Iowa": {"Tier 1": ["Des Moines"], "Tier 2 & 3": ["Cedar Rapids", "Davenport"]},
    "Kansas": {"Tier 1": ["Wichita", "Overland Park"], "Tier 2 & 3": ["Topeka", "Kansas City"]},
    "Kentucky": {"Tier 1": ["Louisville"], "Tier 2 & 3": ["Lexington", "Bowling Green"]},
    "Louisiana": {"Tier 1": ["New Orleans"], "Tier 2 & 3": ["Baton Rouge", "Shreveport", "Lafayette"]},
    "Maine": {"Tier 1": ["Portland"], "Tier 2 & 3": ["Bangor", "Augusta"]},
    "Maryland": {"Tier 1": ["Baltimore"], "Tier 2 & 3": ["Rockville", "Bethesda", "Columbia"]},
    "Massachusetts": {"Tier 1": ["Boston", "Cambridge"], "Tier 2 & 3": ["Worcester", "Springfield"]},
    "Michigan": {"Tier 1": ["Detroit"], "Tier 2 & 3": ["Grand Rapids", "Ann Arbor", "Lansing"]},
    "Minnesota": {"Tier 1": ["Minneapolis", "Saint Paul"], "Tier 2 & 3": ["Rochester", "Bloomington"]},
    "Mississippi": {"Tier 1": ["Jackson"], "Tier 2 & 3": ["Gulfport", "Biloxi"]},
    "Missouri": {"Tier 1": ["St. Louis", "Kansas City"], "Tier 2 & 3": ["Springfield", "Columbia"]},
    "Montana": {"Tier 1": ["Billings"], "Tier 2 & 3": ["Missoula", "Bozeman"]},
    "Nebraska": {"Tier 1": ["Omaha"], "Tier 2 & 3": ["Lincoln", "Bellevue"]},
    "Nevada": {"Tier 1": ["Las Vegas", "Reno"], "Tier 2 & 3": ["Henderson", "Sparks"]},
    "New Hampshire": {"Tier 1": ["Manchester"], "Tier 2 & 3": ["Nashua", "Concord"]},
    "New Jersey": {"Tier 1": ["Jersey City", "Newark"], "Tier 2 & 3": ["Trenton", "Princeton"]},
    "New Mexico": {"Tier 1": ["Albuquerque"], "Tier 2 & 3": ["Santa Fe", "Las Cruces"]},
    "New York": {"Tier 1": ["New York City", "Manhattan", "Brooklyn"], "Tier 2 & 3": ["Buffalo", "Rochester", "Albany", "Syracuse"]},
    "North Carolina": {"Tier 1": ["Charlotte", "Raleigh"], "Tier 2 & 3": ["Durham", "Greensboro", "Winston-Salem"]},
    "North Dakota": {"Tier 1": ["Fargo"], "Tier 2 & 3": ["Bismarck", "Grand Forks"]},
    "Ohio": {"Tier 1": ["Columbus", "Cincinnati", "Cleveland"], "Tier 2 & 3": ["Dayton", "Akron", "Toledo"]},
    "Oklahoma": {"Tier 1": ["Oklahoma City"], "Tier 2 & 3": ["Tulsa", "Norman"]},
    "Oregon": {"Tier 1": ["Portland"], "Tier 2 & 3": ["Eugene", "Salem", "Hillsboro"]},
    "Pennsylvania": {"Tier 1": ["Philadelphia", "Pittsburgh"], "Tier 2 & 3": ["Allentown", "Harrisburg"]},
    "Rhode Island": {"Tier 1": ["Providence"], "Tier 2 & 3": ["Warwick", "Cranston"]},
    "South Carolina": {"Tier 1": ["Charleston", "Columbia"], "Tier 2 & 3": ["Greenville", "Spartanburg"]},
    "South Dakota": {"Tier 1": ["Sioux Falls"], "Tier 2 & 3": ["Rapid City", "Aberdeen"]},
    "Tennessee": {"Tier 1": ["Nashville", "Memphis"], "Tier 2 & 3": ["Knoxville", "Chattanooga"]},
    "Texas": {"Tier 1": ["Austin", "Dallas", "Houston"], "Tier 2 & 3": ["San Antonio", "Fort Worth", "Plano", "Irving"]},
    "Utah": {"Tier 1": ["Salt Lake City"], "Tier 2 & 3": ["Provo", "Lehi", "Ogden"]},
    "Vermont": {"Tier 1": ["Burlington"], "Tier 2 & 3": ["Montpelier", "Rutland"]},
    "Virginia": {"Tier 1": ["Arlington", "Richmond", "Reston"], "Tier 2 & 3": ["Norfolk", "Virginia Beach", "Alexandria"]},
    "Washington": {"Tier 1": ["Seattle", "Bellevue"], "Tier 2 & 3": ["Spokane", "Tacoma", "Redmond"]},
    "West Virginia": {"Tier 1": ["Charleston"], "Tier 2 & 3": ["Huntington", "Morgantown"]},
    "Wisconsin": {"Tier 1": ["Milwaukee", "Madison"], "Tier 2 & 3": ["Green Bay", "Kenosha"]},
    "Wyoming": {"Tier 1": ["Cheyenne"], "Tier 2 & 3": ["Casper", "Laramie"]},
}


def state_dir_name(state):
    return state.lower().replace(" ", "_")


def empty_buckets():

    return {
        it: {scope: [] for scope in PolicyBucketer.SCOPES}
        for it in PolicyBucketer.INCENTIVE_TYPES
    }


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


def _ensure_one(dest, url, referer, force, label):

    if not url:

        if os.path.exists(dest):
            print(f"→ {label} (local file, no URL)")
            return True

        print(f"→ {label}: no URL and local file missing at {dest}")
        return False

    print(f"→ {label}")

    return PolicyDownloader.ensure(url, dest, referer=referer, force=force)


def _download_task(state, filename, url, referer, dest, force):

    label = f"{state} / {filename}"

    ok = _ensure_one(dest, url, referer, force, label)

    return (state, dest if ok else None)


def ensure_pdf_sources(state_pdfs, force=False):

    print(f"Step 1: ensuring USA state PDFs ({DOWNLOAD_WORKERS} workers)\n")

    available = {state: [] for state in state_pdfs}

    tasks = [
        (state, filename, url, referer,
         os.path.join(USA_SOURCES_DIR, state_dir_name(state), filename), force)
        for state, pdfs in state_pdfs.items()
        for filename, url, referer in pdfs
    ]

    with ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as ex:

        futures = [ex.submit(_download_task, *t) for t in tasks]

        for f in as_completed(futures):

            state, path = f.result()

            if path:
                available[state].append(path)

    return available


def _extract_one_state(state, paths, html_urls, state_dir):

    paragraphs = []

    for path in paths:

        try:
            paragraphs.extend(PolicyPDFReader.read(path))
        except Exception:
            pass

    for url in html_urls:

        try:
            paras = PolicyHTMLReader.read(
                url, country_dir=state_dir, depth=1,
                pdf_reader=PolicyPDFReader,
            )
            paragraphs.extend(paras)
        except Exception:
            pass

    if not paragraphs:
        return (state, empty_buckets(), 0)

    return (state, PolicyBucketer.bucket(paragraphs), len(paragraphs))


def build_state_data(state_available, state_html, state_validity, states):

    print(f"\nStep 2: extracting USA states ({EXTRACT_WORKERS} workers)\n")

    state_data = {}
    validity_by_state = {}

    tasks = []

    for state in states:

        validity_by_state[state] = state_validity.get(state, "")

        paths = state_available.get(state, [])
        html_urls = state_html.get(state, [])
        state_dir = os.path.join(USA_SOURCES_DIR, state_dir_name(state))

        tasks.append((state, paths, html_urls, state_dir))

    with ThreadPoolExecutor(max_workers=EXTRACT_WORKERS) as ex:

        futures = {
            ex.submit(_extract_one_state, *t): t[0] for t in tasks
        }

        for f in as_completed(futures):

            state, data, para_count = f.result()

            state_data[state] = data

            total = sum(
                len(items)
                for scopes in data.values()
                for items in scopes.values()
            )

            print(f"→ {state}: bucketed {total} items from {para_count} paragraphs")

    return state_data, validity_by_state


def apply_transform(data, transform_fn, label):

    for jurisdiction, buckets in data.items():

        for incentive_type, scopes in buckets.items():

            for scope, items in scopes.items():

                if not items:
                    continue

                tag = f"{label}/{jurisdiction[:18]}/{incentive_type[:8]}"

                scopes[scope] = transform_fn(items, label=tag)


def build_sources_rows(srcs, state_available):

    rows = []

    for state, pdfs in srcs["country_pdfs"].items():

        downloaded = {os.path.basename(p) for p in state_available.get(state, [])}

        for filename, url, _ in pdfs:

            is_local = not url

            status = "Used" if filename in downloaded else "Failed"

            local = (
                os.path.join(USA_SOURCES_DIR, state_dir_name(state), filename)
                if filename in downloaded else ""
            )

            if is_local:
                note = ("Local file, extracted" if status == "Used"
                        else "Local file missing at expected path")
            else:
                note = ("Downloaded and extracted" if status == "Used"
                        else "Configured but download failed")

            rows.append((
                "USA", state, url or "(local file)", "PDF",
                status, local, note,
            ))

    for state, urls in srcs["country_html"].items():

        for url in urls:

            rows.append((
                "USA", state, url, "HTML", "Scraped", "",
                "HTML scraped and bucketed",
            ))

    return rows


# ---------------------------------------------------------------------------
# Excel export: reuses PolicyExcelExporter's diff/highlight/style building
# blocks (unmodified) rather than its India-only sheet writer, since that
# one hardcodes the "India-State wise" sheet name and region label.
# ---------------------------------------------------------------------------

_PEE = PolicyExcelExporter


def _write_definition_sheet(wb):

    ws = wb.create_sheet("Definition & Methodology")

    rows = [
        ("Definition:",),
        (
            "Incentive Type: The classification of government support offered to Global Capability Centres (GCCs) based on the nature of assistance provided, including the following broad categories:\n"
            "1. Financial Incentives\n2. Infrastructure\n3. Talent\n"
            "4. Regulatory, Ease-of-Doing-Business & Non-financial Support\n"
            "5. Innovation / R&D / Patent\n6. Others",
        ),
        (
            "Applicable in Which Part of the State refers to the geographical scope within a state where a specific incentive is valid. Usage Classification:\n"
            "• Across the state (Common): uniformly applicable across all eligible locations.\n"
            "• Specific to Tier 1 cities: applicable only to designated Tier 1 cities.\n"
            "• Specific to Tier 2 & 3 cities: applicable only to designated Tier 2/3 cities.",
        ),
        (
            "Policy Validity Period: the duration during which a government incentive program remains in force.",
        ),
        ("",),
        ("Methodology:",),
        (
            "US states do not have policies that name Global Capability Centres "
            "directly, unlike India, Mexico, and Poland. This workbook treats "
            "each state's general business investment incentive programs (tax "
            "credits, grants, site selection incentives, R&D credits) as proxy "
            "content, the same approach used for most countries in the Global "
            "sheet of the main GCC incentives workbook.",
        ),
        ("Country Coverage (v1): United States, all 50 states.",),
    ]

    for r, row in enumerate(rows, start=1):

        ws.cell(row=r, column=2, value=row[0]).alignment = _PEE.WRAP

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 120

    for r in range(1, len(rows) + 1):
        ws.row_dimensions[r].height = 60


def _write_usa_sheet(wb, state_data, validity_by_state,
                      baseline=None, has_baseline=False, new_items=None):

    baseline = baseline or {}

    ws = wb.create_sheet("USA-State wise")

    states = list(state_data.keys())

    headers = ["Incentive Type", "Applicable in which part of the state"] + states

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

                items = (
                    state_data.get(state, {})
                    .get(incentive_type, {})
                    .get(scope, [])
                )

                _PEE._write_diff_cell(
                    ws, row_cursor, col_idx, items,
                    baseline.get((row_cursor, col_idx), ""),
                    has_baseline=has_baseline,
                    new_items=new_items,
                    context={
                        "region": "USA",
                        "sheet": "USA-State wise",
                        "jurisdiction": state,
                        "incentive_type": incentive_type,
                        "scope": scope,
                    },
                )

            row_cursor += 1

        if type_start_row + 1 <= row_cursor - 1:

            ws.merge_cells(
                start_row=type_start_row, start_column=1,
                end_row=row_cursor - 1, end_column=1,
            )

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 28

    for col_idx in range(3, 3 + len(states)):

        col_letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[col_letter].width = 55

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30

    for r in range(3, row_cursor):
        ws.row_dimensions[r].height = 220


def export_usa(state_data, validity_by_state, output_path, baseline=None):

    has_baseline = bool(baseline)

    baseline = baseline or {}

    new_items = []

    wb = openpyxl.Workbook()

    wb.remove(wb.active)

    _write_definition_sheet(wb)

    _write_usa_sheet(
        wb, state_data, validity_by_state,
        baseline=baseline.get("USA-State wise", {}),
        has_baseline=has_baseline,
        new_items=new_items,
    )

    _PEE._write_summary_sheet(wb, new_items, has_baseline)

    wb.move_sheet("Summary", offset=-(len(wb.sheetnames) - 2))

    wb.save(output_path)


def output_paths(paraphrased=True):

    today = date.today().isoformat()

    variant = "paraphrased" if paraphrased else "raw"

    return (
        os.path.join(OUTPUT_DIR, f"gcc_incentives_usa_{variant}_{today}.xlsx"),
        os.path.join(OUTPUT_DIR, f"gcc_sources_usa_{variant}_{today}.xlsx"),
    )


# ---------------------------------------------------------------------------
# Flatten step: same 12-column schema as new_template.py, reusing its date
# extraction / city detection / relevance scoring helpers directly so this
# inherits the same accuracy fixes without duplicating that logic.
# ---------------------------------------------------------------------------

def build_usa_source_dates_index(srcs):

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
                USA_SOURCES_DIR, state_dir_name(state), filename), False))

    for state, urls in srcs["country_html"].items():
        for url in urls[:2]:
            if url:
                tasks.append((state, url, True))

    print(f"  ⓘ scanning USA source files for validity dates ({SCAN_WORKERS} workers)")

    def _run(task):
        state, target, is_url = task
        dates = nt.scan_source_for_validity(target, is_url=is_url)
        _remember(state, dates)

    with ThreadPoolExecutor(max_workers=SCAN_WORKERS) as ex:
        list(ex.map(_run, tasks))

    print(f"  ⓘ source date scan found validity for {len(dates_by_state)} states")

    return dates_by_state


def extract_usa_rows_from_workbook(wb, source_dates_index=None):

    if "USA-State wise" not in wb.sheetnames:
        return {}

    ws = wb["USA-State wise"]

    states = []

    for c in range(3, ws.max_column + 1):

        v = ws.cell(1, c).value

        if v:
            states.append((c, str(v).strip()))

    validity_by_col = {
        c: str(ws.cell(2, c).value or "").strip()
        for c, _ in states
    }

    incentive_by_row = nt._extract_incentive_type_map(ws)

    usa_rows_by_state = {state: [] for _, state in states}

    for r in range(3, ws.max_row + 1):

        incentive_type = incentive_by_row.get(r, "")

        scope_label = str(ws.cell(r, 2).value or "").strip()

        if not incentive_type or not scope_label:
            continue

        section = nt.SECTION_MAP.get(incentive_type)

        if not section:
            continue

        scope_type, tier_scope = nt.SCOPE_MAP_INDIA.get(
            scope_label, ("State", scope_label)
        )

        for c, state in states:

            bullets = nt.parse_bullets(ws.cell(r, c).value)

            if not bullets:
                continue

            validity_start, vs_prec, validity_end, ve_prec = nt.parse_validity(
                validity_by_col.get(c, "")
            )

            for b in bullets:

                text_dates = nt.extract_validity_from_text(b)

                start, s_prec, end, e_prec = nt._apply_source_fallback(
                    text_dates,
                    (source_dates_index or {}).get(state),
                    (validity_start, vs_prec, validity_end, ve_prec),
                )

                if nt.is_expired(end, e_prec):
                    continue

                cities = nt.detect_cities(b, US_STATE_CITIES.get(state, {}))

                relevance = nt.compute_gcc_relevance(b, "United States")

                if cities:
                    for city_name, tier in cities:
                        usa_rows_by_state[state].append({
                            "section_name": section,
                            "scope_type": "City",
                            "country_name": "United States",
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
                    usa_rows_by_state[state].append({
                        "section_name": section,
                        "scope_type": scope_type,
                        "country_name": "United States",
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

    return usa_rows_by_state


def write_new_template_usa(usa_rows_by_state, output_path):

    all_rows = []

    for state in usa_rows_by_state:

        rows = sorted(
            usa_rows_by_state[state],
            key=lambda r: -(r.get("gcc_relevance") or 0),
        )

        all_rows.extend(rows)

    signal_rows = [r for r in all_rows if int(r.get("gcc_relevance") or 0) >= 1]
    other_rows = [r for r in all_rows if int(r.get("gcc_relevance") or 0) == 0]

    wb = openpyxl.Workbook()

    wb.remove(wb.active)

    nt._write_flat_sheet(wb, "United States", signal_rows)
    nt._write_flat_sheet(wb, "Other - United States", other_rows)

    wb.save(output_path)

    return len(signal_rows), len(other_rows)


def flatten_latest(paraphrase_preferred=True):
    """Read the most recently written USA workbook and produce the flat
    new_template style output. Mirrors new_template.run(), scoped to USA."""

    candidates = [
        f for f in os.listdir(OUTPUT_DIR)
        if f.startswith("gcc_incentives_usa_") and f.endswith(".xlsx")
        and not f.startswith("~$")
    ] if os.path.isdir(OUTPUT_DIR) else []

    if not candidates:
        print("✗ No gcc_incentives_usa_*.xlsx found in output/United States/.")
        return None

    def _sort_key(name):
        full = os.path.join(OUTPUT_DIR, name)
        prefers = 1 if (paraphrase_preferred and "paraphrased" in name) else 0
        return (prefers, os.path.getmtime(full))

    latest = max(candidates, key=_sort_key)
    latest_path = os.path.join(OUTPUT_DIR, latest)

    print(f"→ Reading latest USA workbook: {latest}")

    wb = openpyxl.load_workbook(latest_path, data_only=True)

    source_dates_index = {}

    if os.path.exists(SOURCES_FILE):

        try:
            srcs = SourcesLoader.load(SOURCES_FILE)
            source_dates_index = build_usa_source_dates_index(srcs)
        except Exception as exc:
            print(f"  ⚠ source date scan skipped: {exc}")

    usa_by_state = extract_usa_rows_from_workbook(wb, source_dates_index=source_dates_index)

    today = date.today().isoformat()

    out_path = os.path.join(OUTPUT_DIR, f"new_template_USA_policy_{today}.xlsx")

    signal_count, other_count = write_new_template_usa(usa_by_state, out_path)

    print(f"\n✓ Wrote {out_path}")
    print(f"  United States rows: {signal_count} (score >= 1)")
    print(f"  Other United States: {other_count} (score = 0)")

    return out_path


def main():

    parser = argparse.ArgumentParser(
        description="Build the USA GCC incentives workbook from input/sources_usa.xlsx."
    )

    parser.add_argument(
        "--refresh", action="store_true",
        help="Force re-download of all PDFs, ignoring local cache.",
    )

    parser.add_argument(
        "--sources", default=SOURCES_FILE,
        help=f"Path to USA sources workbook (default: {SOURCES_FILE})",
    )

    parser.add_argument(
        "--nopara", "--no-paraphrase", dest="paraphrase",
        action="store_false", default=True,
        help="Skip the Ollama paraphrasing step.",
    )

    parser.add_argument(
        "--notranslate", "--no-translate", dest="translate",
        action="store_false", default=True,
        help="Skip the argos-translate step.",
    )

    parser.add_argument(
        "--paraphrase", dest="paraphrase", action="store_true",
        help=argparse.SUPPRESS,
    )

    parser.add_argument(
        "--model", default=OllamaParaphraser.DEFAULT_MODEL,
        help=f"Ollama model name for paraphrasing (default: {OllamaParaphraser.DEFAULT_MODEL})",
    )

    args = parser.parse_args()

    print(f"Loading USA sources from {args.sources}\n")

    srcs = SourcesLoader.load(args.sources)

    states_from_sources = _ordered_keys(
        srcs["country_pdfs"], srcs["country_html"], srcs["country_validity"]
    )

    states = _merge_lists(US_STATES, states_from_sources)

    print(f"  → USA states: {len(states)}\n")

    os.makedirs(USA_SOURCES_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    state_available = ensure_pdf_sources(srcs["country_pdfs"], force=args.refresh)

    state_data, validity_by_state = build_state_data(
        state_available, srcs["country_html"], srcs["country_validity"], states,
    )

    if args.translate:

        print("\nStep 2b: translating non-English text via argos-translate\n")

        tr = ArgosTranslator()

        if tr.is_available():
            apply_transform(state_data, tr.translate_many, "states")

    if args.paraphrase:

        print("\nStep 2c: paraphrasing via Ollama\n")

        para = OllamaParaphraser(model=args.model)

        if not para.is_available():

            print(
                "⚠ Ollama not reachable at "
                f"{OllamaParaphraser.HOST} or model '{args.model}' not pulled. "
                f"Skipping paraphrasing. Run: ollama pull {args.model}"
            )

        else:
            apply_transform(state_data, para.paraphrase_many, "states")

    print("\nStep 3: writing workbooks\n")

    out_path, sources_path = output_paths(paraphrased=args.paraphrase)

    baseline_prefix = (
        "gcc_incentives_usa_paraphrased_" if args.paraphrase
        else "gcc_incentives_usa_raw_"
    )

    baseline_path = PreviousWorkbookReader.find_latest(
        OUTPUT_DIR, baseline_prefix, exclude_path=out_path,
    )

    baseline = PreviousWorkbookReader.load(baseline_path)

    if baseline:
        print(f"  ⟂ comparing against baseline: {os.path.basename(baseline_path)}")
    else:
        print("  ⟂ no prior USA workbook found; all content rendered black")

    export_usa(state_data, validity_by_state, out_path, baseline=baseline)

    print(f"  ✓ Wrote {out_path}")

    rows = build_sources_rows(srcs, state_available)

    SourcesExcelExporter.export(rows, sources_path)

    print(f"  ✓ Wrote {sources_path}")

    print("\nStep 4: building new_template flat workbook for USA\n")

    try:
        flatten_latest(paraphrase_preferred=args.paraphrase)
    except Exception as exc:
        print(f"⚠ flatten_latest() failed: {exc}. "
              "You can re-run 'python3 main_usa.py' to retry the flatten step.")


if __name__ == "__main__":
    main()
