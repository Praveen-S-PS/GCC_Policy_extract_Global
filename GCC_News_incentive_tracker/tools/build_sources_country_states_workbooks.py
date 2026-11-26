"""One-time generator for input/sources_<country>.xlsx (United Kingdom,
Germany, France, Japan), parsing the research table accumulated in
country_states_research.md into the standard 8-column sources schema.

Does not touch input/sources.xlsx, input/sources_usa.xlsx, or any other
existing sources file.
"""

import os
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


RESEARCH_FILE = (
    "/private/tmp/claude-502/-Users-test-Desktop-Claude/"
    "6c9e54d3-ea51-4f29-b355-67f197eca528/scratchpad/country_states_research.md"
)

COUNTRY_LABEL_MAP = {
    "Japan": "Japan",
    "UK": "United Kingdom",
    "Germany": "Germany",
    "France": "France",
}

# Extra rows not present in the parsed research file (found in an earlier,
# separate research pass).
EXTRA_ROWS = {
    "Germany": [
        ("Saxony-Anhalt", "IMG Sachsen-Anhalt",
         "Support and financing service",
         "https://www.invest-in-saxony-anhalt.com/investor-services/support-and-financing-service",
         "HTML", "English", "confirmed (already verified in an earlier pass)"),
    ],
}

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
HEADERS = ["Active", "Region", "Jurisdiction", "Filename", "Validity", "Source URL", "Referer", "Notes"]


def slug(name):
    return name.lower().replace(" ", "_").replace("-", "_").replace("'", "")


def parse_research_file(path):

    rows_by_country = {c: [] for c in COUNTRY_LABEL_MAP.values()}

    with open(path, encoding="utf-8") as f:
        lines = f.readlines()

    for line in lines:

        line = line.strip()

        if not line or line.startswith("#") or line.startswith("Format:"):
            continue

        parts = [p.strip() for p in line.split("|")]

        if len(parts) < 7:
            continue

        raw_country = parts[0]

        if raw_country not in COUNTRY_LABEL_MAP:
            continue

        country = COUNTRY_LABEL_MAP[raw_country]

        jurisdiction, agency, title, url, fmt, language = parts[1:7]
        notes = parts[7] if len(parts) > 7 else ""

        if not url.startswith("http"):
            continue

        rows_by_country[country].append(
            (jurisdiction, agency, title, url, fmt, language, notes)
        )

    for country, extra in EXTRA_ROWS.items():
        rows_by_country[country].extend(extra)

    return rows_by_country


def write_workbook(country, rows, dest):

    wb = Workbook()
    ws = wb.active
    ws.title = "Sources"

    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER

    for r_idx, (jurisdiction, agency, title, url, fmt, language, notes) in enumerate(rows, start=2):

        note = f"{agency} — {title} ({language}). {notes}".strip()

        values = ["yes", "Global", jurisdiction, "", "", url, "", note]

        for c_idx, v in enumerate(values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = BORDER
            if c_idx == 6 and isinstance(v, str) and v.startswith("http"):
                cell.hyperlink = v
                cell.font = Font(color="0563C1", underline="single")

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 9
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 75
    ws.column_dimensions["G"].width = 10
    ws.column_dimensions["H"].width = 70

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28

    for r in range(2, len(rows) + 2):
        ws.row_dimensions[r].height = 45

    os.makedirs("input", exist_ok=True)
    wb.save(dest)

    print(f"Wrote {dest} with {len(rows)} jurisdiction rows")


def main():

    rows_by_country = parse_research_file(RESEARCH_FILE)

    for country, rows in rows_by_country.items():

        dest = os.path.join("input", f"sources_{slug(country)}.xlsx")

        write_workbook(country, rows, dest)


if __name__ == "__main__":
    main()
