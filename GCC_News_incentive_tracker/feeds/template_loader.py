import os

import openpyxl


class TemplateLoader:
    """Reads input/template.xlsx to get the list of country sheets to include."""

    COUNTRIES_SHEET = "Countries"

    @classmethod
    def load_countries(cls, template_path):

        if not os.path.exists(template_path):
            return []

        try:

            wb = openpyxl.load_workbook(template_path, data_only=True)

        except Exception:

            return []

        if cls.COUNTRIES_SHEET not in wb.sheetnames:
            return []

        ws = wb[cls.COUNTRIES_SHEET]

        countries = []

        seen = set()

        for r in range(2, ws.max_row + 1):

            name = ws.cell(row=r, column=1).value

            if not name or not isinstance(name, str):
                continue

            name = name.strip()

            if not name or name in seen:
                continue

            seen.add(name)

            countries.append(name)

        return countries
