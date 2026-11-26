"""One time generator for input/template.xlsx.

Creates a 4 sheet reference/config workbook:
  1. Definition & Methodology   — same content as expected workbook
  2. India-State wise           — blank schema (8 state columns)
  3. Country Template           — blank generic schema (single "Country" column)
  4. Countries                  — single column list of country names to include

Users edit the Countries sheet to add or remove country sheets from the output
workbook. The pipeline merges this list with jurisdictions in sources.xlsx.
"""

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


DEST = "input/template.xlsx"


HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TYPE_FILL = PatternFill("solid", fgColor="D9E1F2")
BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")


INCENTIVE_TYPES = [
    "Financial Incentives",
    "Infrastructure",
    "Talent",
    "Regulatory, Ease-of-Doing-Business & Non-financial Support",
    "Innovation / R&D / Patent",
    "Others",
]

INDIA_SCOPES = [
    "Across the state (Common)",
    "Specific to Tier 1 cities",
    "Specific to Tier 2 & 3 cities",
]

COUNTRY_SCOPES = [
    "Across the Country (Common)",
    "Specific to Tier 1 cities",
    "Specific to Tier 2 & 3 cities",
]

INDIA_STATES = [
    "Gujarat", "Karnataka", "Madhya Pradesh", "Maharashtra",
    "Rajasthan", "Tamil Nadu", "Telangana", "Uttar Pradesh",
]

DEFAULT_COUNTRIES = [
    "Albania", "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan",
    "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium",
    "Belize", "Bosnia and Herzegovina", "Brazil", "Brunei", "Bulgaria",
    "Cambodia", "Canada", "Chile", "China", "Colombia", "Costa Rica",
    "Cote d'Ivoire", "Croatia", "Cuba", "Cyprus", "Czech Republic", "Denmark",
    "Dominican Republic", "Ecuador", "Egypt", "El Salvador", "Estonia",
    "Ethiopia", "Fiji", "Finland", "France", "Georgia", "Germany", "Ghana",
    "Greece", "Guatemala", "Guyana", "Honduras", "Hungary", "Iceland",
    "India", "Indonesia", "Ireland", "Israel", "Italy", "Jamaica", "Japan",
    "Jordan", "Kenya", "Kuwait", "Laos", "Latvia", "Lithuania", "Luxembourg",
    "Malaysia", "Malta", "Mauritius", "Mexico", "Montenegro", "Morocco",
    "Netherlands", "New Zealand", "Nicaragua", "Nigeria", "North Macedonia",
    "Norway", "Oman", "Panama", "Paraguay", "Peru", "Philippines", "Poland",
    "Portugal", "Puerto Rico", "Qatar", "Romania", "Rwanda", "Saudi Arabia",
    "Senegal", "Serbia", "Singapore", "Slovakia", "Slovenia", "South Africa",
    "South Korea", "Spain", "Sri Lanka", "Sweden", "Switzerland", "Taiwan",
    "Thailand", "Trinidad and Tobago", "Tunisia", "Turkey", "UAE", "Uganda",
    "United Kingdom", "United States", "Uruguay", "Uzbekistan", "Vietnam",
    "Zambia", "Zimbabwe",
]


def write_definition_sheet(wb):

    ws = wb.create_sheet("Definition & Methodology")

    rows = [
        "Definition:",
        (
            "Incentive Type: The classification of government support offered to "
            "Global Capability Centres (GCCs) based on the nature of assistance:\n"
            "1. Financial Incentives\n2. Infrastructure\n3. Talent\n"
            "4. Regulatory, Ease-of-Doing-Business & Non-financial Support\n"
            "5. Innovation / R&D / Patent\n6. Others"
        ),
        (
            "Applicable in Which Part of the State/Country: geographic scope of "
            "the incentive.\n"
            "• Across the state/country (Common): applies everywhere.\n"
            "• Specific to Tier 1 cities.\n"
            "• Specific to Tier 2 & 3 cities."
        ),
        "Policy Validity Period: the duration during which a policy remains in force.",
        (
            "Caveat/Conditions: eligibility criteria, time limits, compliance "
            "obligations, exclusions, and performance-linked requirements."
        ),
        "",
        "Methodology:",
        (
            "This workbook is built by an automated pipeline that reads policy "
            "PDFs and government landing pages from sources listed in "
            "input/sources.xlsx, extracts text, classifies each clause into the "
            "6 incentive types by 3 scope tiers, and paraphrases via a local "
            "LLM (Ollama). New content since the previous run is highlighted in "
            "red font, with an at-a-glance change list on the Summary sheet."
        ),
        (
            "Country Coverage: driven by the Countries sheet in this template "
            "and by jurisdictions in input/sources.xlsx. To add a country, "
            "either add its name to the Countries sheet (empty placeholder "
            "sheet appears) or add rows to sources.xlsx with its source URLs "
            "(populated sheet appears)."
        ),
    ]

    for r, text in enumerate(rows, start=1):

        cell = ws.cell(row=r, column=2, value=text)
        cell.alignment = WRAP

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 120

    for r in range(1, len(rows) + 1):
        ws.row_dimensions[r].height = 60


def write_schema_sheet(ws, scope_list, jurisdiction_columns, scope_label):
    """Write the 6 x 3 empty schema."""

    headers = ["Incentive Type", scope_label] + jurisdiction_columns

    for c, h in enumerate(headers, start=1):

        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(
            wrap_text=True, vertical="center", horizontal="center"
        )
        cell.border = BORDER

    cell = ws.cell(row=2, column=1, value="Policy Validity Period ->")
    cell.font = Font(bold=True)
    cell.alignment = WRAP
    cell.border = BORDER
    ws.cell(row=2, column=2).border = BORDER

    for col_idx in range(3, 3 + len(jurisdiction_columns)):
        ws.cell(row=2, column=col_idx).border = BORDER
        ws.cell(row=2, column=col_idx).alignment = WRAP

    row_cursor = 3

    for incentive_type in INCENTIVE_TYPES:

        type_start_row = row_cursor

        for idx, scope in enumerate(scope_list):

            type_cell = ws.cell(row=row_cursor, column=1)

            if idx == 0:
                type_cell.value = incentive_type
                type_cell.font = Font(bold=True)
                type_cell.fill = TYPE_FILL

            type_cell.alignment = WRAP
            type_cell.border = BORDER

            scope_cell = ws.cell(row=row_cursor, column=2, value=scope)
            scope_cell.alignment = WRAP
            scope_cell.border = BORDER

            for col_idx in range(3, 3 + len(jurisdiction_columns)):
                ws.cell(row=row_cursor, column=col_idx).border = BORDER
                ws.cell(row=row_cursor, column=col_idx).alignment = WRAP

            row_cursor += 1

        if type_start_row + 1 <= row_cursor - 1:
            ws.merge_cells(
                start_row=type_start_row, start_column=1,
                end_row=row_cursor - 1, end_column=1,
            )

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 28

    for col_idx in range(3, 3 + len(jurisdiction_columns)):
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[col_letter].width = 55

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 30

    for r in range(3, row_cursor):
        ws.row_dimensions[r].height = 60


def write_india_sheet(wb):

    ws = wb.create_sheet("India-State wise")

    write_schema_sheet(
        ws, INDIA_SCOPES, INDIA_STATES,
        "Applicable in which part of the state",
    )


def write_country_template_sheet(wb):

    ws = wb.create_sheet("Country Template")

    write_schema_sheet(
        ws, COUNTRY_SCOPES, ["<Country Name>"],
        "Applicable in which part of the Country",
    )


def write_countries_list_sheet(wb):

    ws = wb.create_sheet("Countries")

    ws.cell(row=1, column=1, value="Countries").font = HEADER_FONT
    ws.cell(row=1, column=1).fill = HEADER_FILL
    ws.cell(row=1, column=1).alignment = Alignment(
        wrap_text=True, vertical="center", horizontal="center"
    )
    ws.cell(row=1, column=1).border = BORDER

    ws.cell(row=1, column=2, value="Notes").font = HEADER_FONT
    ws.cell(row=1, column=2).fill = HEADER_FILL
    ws.cell(row=1, column=2).alignment = Alignment(
        wrap_text=True, vertical="center", horizontal="center"
    )
    ws.cell(row=1, column=2).border = BORDER

    for r_idx, country in enumerate(DEFAULT_COUNTRIES, start=2):

        cell = ws.cell(row=r_idx, column=1, value=country)
        cell.alignment = WRAP
        cell.border = BORDER

        ws.cell(row=r_idx, column=2).border = BORDER

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 60
    ws.row_dimensions[1].height = 30

    for r in range(2, len(DEFAULT_COUNTRIES) + 2):
        ws.row_dimensions[r].height = 22


def main():

    os.makedirs("input", exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)

    write_definition_sheet(wb)
    write_india_sheet(wb)
    write_country_template_sheet(wb)
    write_countries_list_sheet(wb)

    wb.save(DEST)

    print(f"Wrote {DEST}")

    print(
        "Sheets:",
        ", ".join(wb.sheetnames),
    )


if __name__ == "__main__":
    main()
