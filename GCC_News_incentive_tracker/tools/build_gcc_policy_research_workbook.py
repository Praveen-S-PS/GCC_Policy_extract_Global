"""Shared writer for the "does this country/state have a dedicated GCC
policy" research workbooks (one file per country, per the user's request).

Each call to build_workbook() produces a single .xlsx with:
  - a Definition & Methodology sheet explaining what (a)/(b)/(c) mean and
    how the search was run (full phrase "Global Capability Center/Centre",
    never the bare "GCC" acronym, to avoid Gulf Cooperation Council noise)
  - a Findings sheet: one row per jurisdiction checked

Does not read or write any of the main pipeline's input/output files.
"""

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")

RESULT_LABELS = {
    "a": "(a) Dedicated GCC-named policy found",
    "b": "(b) Mentioned as an eligible/targeted category",
    "c": "(c) No mention found",
}

FINDINGS_HEADERS = [
    "Jurisdiction", "Result", "Result Meaning", "Source URL", "Quoted Phrase",
]


def build_workbook(country, coverage_note, rows, summary, output_path):
    """rows: list of (jurisdiction, result_code, url, quote) tuples."""

    wb = Workbook()

    wb.remove(wb.active)

    _write_methodology_sheet(wb, country, coverage_note, summary)
    _write_findings_sheet(wb, rows)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    wb.save(output_path)

    print(f"Wrote {output_path} with {len(rows)} jurisdiction rows")


def _write_methodology_sheet(wb, country, coverage_note, summary):

    ws = wb.create_sheet("Definition & Methodology")

    rows = [
        (f"GCC Policy Research: {country}",),
        ("",),
        (
            "Question: does this country, or any of its internal states/"
            "regions/prefectures, have a government policy that explicitly "
            "names \"Global Capability Centre(s)\" or \"Global Capability "
            "Center(s)\" as its subject or an eligible category -- the same "
            "way several Indian states have a standalone \"GCC Policy\" "
            "document?",
        ),
        (
            "Method: searched for the literal phrase \"Global Capability "
            "Center\" / \"Global Capability Centre\" only. The bare "
            "acronym \"GCC\" was deliberately never used as a search term, "
            "since it collides heavily with the Gulf Cooperation Council "
            "and other unrelated organizations and would return unusable "
            "noise.",
        ),
        (
            "Result codes:\n"
            "(a) Dedicated policy -- an official government source names "
            "\"Global Capability Centre(s)\" as the actual subject of a "
            "policy, program, or tax credit.\n"
            "(b) Mentioned as a category -- an official government or "
            "investment-promotion-agency source mentions the phrase as one "
            "eligible/targeted category among others, without a dedicated "
            "policy built around it.\n"
            "(c) No mention -- no official government or agency source "
            "uses the phrase at all.",
        ),
        (
            "Important exclusion: a private company's own press release, "
            "or a news article reporting that some company opened a "
            "facility it calls a \"Global Capability Centre,\" does NOT "
            "count toward (a) or (b). Only the government's own policy/"
            "program language counts.",
        ),
        ("",),
        (f"Coverage note: {coverage_note}",),
        ("",),
        (f"Summary: {summary}",),
    ]

    for r, row in enumerate(rows, start=1):

        ws.cell(row=r, column=2, value=row[0]).alignment = WRAP

    ws.cell(row=1, column=2).font = Font(bold=True, size=14)

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 110

    for r in range(1, len(rows) + 1):
        ws.row_dimensions[r].height = 60

    ws.row_dimensions[1].height = 30


def _write_findings_sheet(wb, rows):

    ws = wb.create_sheet("Findings")

    for c, h in enumerate(FINDINGS_HEADERS, start=1):

        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(
            wrap_text=True, vertical="center", horizontal="center"
        )
        cell.border = BORDER

    for r_idx, (jurisdiction, code, url, quote) in enumerate(rows, start=2):

        values = [
            jurisdiction,
            code,
            RESULT_LABELS.get(code, code),
            url or "",
            quote or "none found",
        ]

        for c_idx, v in enumerate(values, start=1):

            cell = ws.cell(row=r_idx, column=c_idx, value=v)
            cell.alignment = WRAP
            cell.border = BORDER

            if c_idx == 4 and isinstance(v, str) and v.startswith("http"):
                cell.hyperlink = v
                cell.font = Font(color="0563C1", underline="single")

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 42
    ws.column_dimensions["D"].width = 55
    ws.column_dimensions["E"].width = 70

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28

    for r in range(2, len(rows) + 2):
        ws.row_dimensions[r].height = 40
