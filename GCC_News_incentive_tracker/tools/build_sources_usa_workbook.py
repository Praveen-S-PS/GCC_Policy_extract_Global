"""One time generator for input/sources_usa.xlsx.

Writes one row per US state (plus an optional District of Columbia row) into
a single editable Excel file with the same schema as input/sources.xlsx.
After running this once, edit input/sources_usa.xlsx directly to add,
remove, or correct a state's source without touching code.

Every URL below was found and fetch-verified via web research (see the
research notes column) rather than guessed. A few states did not have a
single clean official source available (a domain blocked automated
fetching, or the agency's own site was reorganized); those are flagged in
the Notes column as "best effort" and are good candidates for a human to
double check or replace.

Does not read or write input/sources.xlsx or input/template.xlsx.
"""

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


DEST = "input/sources_usa.xlsx"


# (state, filename_or_None, validity_or_None, url, referer_or_None, notes)
# filename is only needed for PDF sources; HTML sources are scraped directly.
PRECONFIGURED_USA = [
    ("Alabama", None, None,
     "https://www.madeinalabama.com/why-alabama/taxes-incentives/", None,
     "Alabama Dept of Commerce (Made in Alabama) — Tax Incentives and Grants"),
    ("Alaska", None, None,
     "https://www.tax.alaska.gov/programs/programs/credits/index.aspx", None,
     "Best effort: AK DCCED site blocked by bot protection; substituted AK Dept of Revenue Tax Credits page"),
    ("Arizona", None, None,
     "https://www.azcommerce.com/programs/", None,
     "Arizona Commerce Authority — Programs for Businesses"),
    ("Arkansas", None, None,
     "https://www.arkansasedc.com/why-arkansas/business-climate/incentives", None,
     "Arkansas Economic Development Commission — Incentives"),
    ("California", None, None,
     "https://business.ca.gov/resources/incentives-grants-and-financing/calbis-resources/california-business-incentives-list/", None,
     "Governor's Office of Business and Economic Development (GO-Biz) — CA Business Incentives List"),
    ("Colorado", None, None,
     "https://oedit.colorado.gov/business-funding-and-incentives", None,
     "Colorado OEDIT — Business Funding and Incentives"),
    ("Connecticut", None, None,
     "https://portal.ct.gov/decd/services/business-development/tax-incentives", None,
     "Connecticut DECD — Tax Incentives"),
    ("Delaware", None, None,
     "https://business.delaware.gov/incentives", None,
     "Delaware Division of Small Business — Incentives & Credits"),
    ("Florida", None, None,
     "https://floridarevenue.com/taxes/taxesfees/Pages/corp_tax_incent.aspx", None,
     "Best effort: Enterprise Florida dissolved; substituted FL Dept of Revenue Corporate Income Tax Incentives page"),
    ("Georgia", None, None,
     "https://georgia.org/competitive-advantages/incentives/tax-credits", None,
     "Georgia Dept of Economic Development — Georgia's Incentives"),
    ("Hawaii", None, None,
     "https://invest.hawaii.gov/business-programs/business-incentives-and-programs/", None,
     "Hawaii DBEDT — Business Incentives and Programs"),
    ("Idaho", None, None,
     "https://commerce.idaho.gov/incentives-and-financing/incentives/", None,
     "Idaho Commerce — Tax Incentives"),
    ("Illinois", None, None,
     "https://dceo.illinois.gov/businesshelp/incentivesandtaxcredits.html", None,
     "Illinois DCEO — Incentives and Tax Credits"),
    ("Indiana", None, None,
     "https://www.iedc.in.gov/indiana-advantages/investments", None,
     "Indiana Economic Development Corporation — Grants & Incentives"),
    ("Iowa", None, None,
     "https://opportunityiowa.gov/business/financial-assistance", None,
     "Iowa Economic Development Authority — Financial Assistance"),
    ("Kansas", None, None,
     "https://www.kansascommerce.gov/businesses/incentives/", None,
     "Kansas Dept of Commerce — Business Incentives and Financing"),
    ("Kentucky", None, None,
     "https://newkentuckyhome.ky.gov/Locating_Expanding/Financial_Incentives", None,
     "Kentucky Cabinet for Economic Development — Financial Incentives & Programs"),
    ("Louisiana", None, None,
     "https://www.opportunitylouisiana.gov/incentives", None,
     "Louisiana Economic Development — Business Incentives"),
    ("Maine", None, None,
     "https://www.maine.gov/decd/business-development/financial-incentives-resources/incentives", None,
     "Maine DECD — Financial Incentives"),
    ("Maryland", "maryland_financial_incentives_brochure.pdf", None,
     "https://commerce.maryland.gov/Documents/BusinessResource/maryland-financial-incentives-for-businesses-brochure.pdf", None,
     "Best effort: official brochure PDF, but document metadata dates it to 2018"),
    ("Massachusetts", None, None,
     "https://www.mass.gov/info-details/economic-development-incentive-program-edip", None,
     "Best effort: mass.gov blocks automated fetch; verified via archived snapshot. Mass Office of Business Development — EDIP"),
    ("Michigan", None, None,
     "https://www.michiganbusiness.org/services/incentives-and-taxes/", None,
     "Michigan Economic Development Corporation — Incentives and Taxes"),
    ("Minnesota", None, None,
     "https://mn.gov/deed/business/locating-minnesota/incentives/", None,
     "Minnesota DEED — Incentives"),
    ("Mississippi", None, None,
     "https://mississippi.org/doing-business/incentives/", None,
     "Mississippi Development Authority — Incentives Directory"),
    ("Missouri", None, None,
     "https://missouripartnership.com/why-missouri/incentive-programs/", None,
     "Best effort: ded.mo.gov blocked all fetch attempts; substituted the state-authorized Missouri Partnership site"),
    ("Montana", None, None,
     "https://commerce.mt.gov/Business/Business-Assistance/", None,
     "Best effort: Montana Dept of Commerce — Business Assistance (leans grants/loans over tax credits)"),
    ("Nebraska", None, None,
     "https://opportunity.nebraska.gov/programs/incentives/", None,
     "Nebraska Dept of Economic Development — Incentive Programs"),
    ("Nevada", None, None,
     "https://www.expand2nevada.com/programs-resources/incentives", None,
     "Nevada Governor's Office of Economic Development (GOED) — Incentives"),
    ("New Hampshire", None, None,
     "https://www.nheconomy.com/why-new-hampshire/incentives-taxes", None,
     "NH Dept of Business and Economic Affairs — Incentives & Taxes"),
    ("New Jersey", None, None,
     "https://www.njeda.gov/financing-and-incentives/", None,
     "New Jersey Economic Development Authority — Financing and Incentives"),
    ("New Mexico", None, None,
     "https://edd.newmexico.gov/choose-new-mexico/competitive-business-climate/", None,
     "New Mexico Economic Development Dept — Competitive Business Climate"),
    ("New York", None, None,
     "https://www.esd.ny.gov/tax-based-incentives", None,
     "Empire State Development — New York Tax-Based Incentives"),
    ("North Carolina", None, None,
     "https://www.commerce.nc.gov/grants-incentives", None,
     "North Carolina Dept of Commerce — Grants & Incentives"),
    ("North Dakota", None, None,
     "https://www.commerce.nd.gov/economic-development-finance/finance-and-incentives", None,
     "North Dakota Dept of Commerce — Finance and Incentives"),
    ("Ohio", None, None,
     "https://www.jobsohio.com/incentives-programs", None,
     "JobsOhio — Explore Ohio's Business Tax Incentives, Grants, & Loans"),
    ("Oklahoma", None, None,
     "https://www.okcommerce.gov/doing-business/business-relocation-expansion/incentives/", None,
     "Oklahoma Dept of Commerce — Business Incentives"),
    ("Oregon", None, None,
     "https://www.oregon.gov/biz/programs/business_incentives/pages/default.aspx", None,
     "Business Oregon — Business Incentives"),
    ("Pennsylvania", None, None,
     "https://dced.pa.gov/program/", None,
     "PA Dept of Community and Economic Development — Programs and Funding"),
    ("Rhode Island", None, None,
     "https://commerceri.com/financing/tax-credits/", None,
     "Rhode Island Commerce Corporation — Tax Credits"),
    ("South Carolina", "south_carolina_business_incentives_booklet.pdf", None,
     "https://s3.amazonaws.com/cdn.sccommerce.com/scdoc-public/2026-02/SCDofC-BusinessIncentives-Booklet_January%202026.pdf", None,
     "SC Dept of Commerce — Business Incentives Booklet (January 2026 edition)"),
    ("South Dakota", None, None,
     "https://sdgoed.com/financing-incentives/", None,
     "South Dakota Governor's Office of Economic Development — Financing + Incentives"),
    ("Tennessee", None, None,
     "https://tnecd.com/advantages/incentives-grants/", None,
     "Tennessee Dept of Economic and Community Development (TNECD) — Incentives & Grants"),
    ("Texas", None, None,
     "https://gov.texas.gov/business/page/incentives", None,
     "Texas Economic Development & Tourism Office — Incentive & Financial Programs"),
    ("Utah", None, None,
     "https://business.utah.gov/recruitment/", None,
     "Governor's Office of Economic Opportunity — Business Recruitment & Expansion"),
    ("Vermont", None, None,
     "https://accd.vermont.gov/economic-development/business-resources", None,
     "Vermont Agency of Commerce and Community Development — Business Resources"),
    ("Virginia", None, None,
     "https://www.vedp.org/incentives", None,
     "Virginia Economic Development Partnership — Incentives"),
    ("Washington", None, None,
     "https://choosewashingtonstate.com/i-need-help-with/site-selection/incentives/", None,
     "WA Dept of Commerce (ChooseWashington) — WA State Business Tax Incentives"),
    ("West Virginia", None, None,
     "https://westvirginia.gov/wv-incentives/", None,
     "West Virginia Division of Economic Development — WV Incentives"),
    ("Wisconsin", "wisconsin_wedc_annual_report_fy25.pdf", None,
     "https://wedc.org/wp-content/uploads/2025/10/WEDC-ARED-FY25-100125.pdf", None,
     "Wisconsin Economic Development Corporation — Annual Report on Economic Development FY2025"),
    ("Wyoming", None, None,
     "https://wyomingbusiness.org/business/financing/financial-incentives/", None,
     "Best effort: Wyoming Business Council — Financial Incentives (data center incentives live on separate pages)"),
]

# Not one of the 50 states -- included only because it has a genuinely
# notable, GCC-relevant program. Leave Active = "no" by default so it does
# not show up in a run unless a human deliberately flips it on.
PRECONFIGURED_DC = [
    ("District of Columbia", None, None,
     "https://dmped.dc.gov/qhtc", None,
     "Not one of the 50 states -- DC Deputy Mayor for Planning and Economic Development, Qualified High Tech Companies (QHTC) program. Active=no by default."),
]


HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

HEADERS = [
    "Active",
    "Region",
    "Jurisdiction",
    "Filename",
    "Validity",
    "Source URL",
    "Referer",
    "Notes",
]


def main():

    rows = []

    for state, filename, validity, url, referer, notes in PRECONFIGURED_USA:

        rows.append([
            "yes",
            "USA",
            state,
            filename or "",
            validity or "",
            url,
            referer or "",
            notes,
        ])

    for state, filename, validity, url, referer, notes in PRECONFIGURED_DC:

        rows.append([
            "no",
            "USA",
            state,
            filename or "",
            validity or "",
            url,
            referer or "",
            notes,
        ])

    wb = Workbook()

    ws = wb.active

    ws.title = "Sources"

    for c, h in enumerate(HEADERS, start=1):

        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(
            wrap_text=True, vertical="center", horizontal="center"
        )
        cell.border = BORDER

    for r_idx, row in enumerate(rows, start=2):

        for c_idx, value in enumerate(row, start=1):

            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(
                wrap_text=True, vertical="top", horizontal="left"
            )
            cell.border = BORDER

            if c_idx == 6 and isinstance(value, str) and value.startswith("http"):

                cell.hyperlink = value
                cell.font = Font(color="0563C1", underline="single")

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 9
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 32
    ws.column_dimensions["E"].width = 20
    ws.column_dimensions["F"].width = 70
    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["H"].width = 55

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30

    for r in range(2, len(rows) + 2):
        ws.row_dimensions[r].height = 30

    os.makedirs("input", exist_ok=True)

    wb.save(DEST)

    print(f"Wrote {DEST} with {len(rows)} source rows "
          f"({len(PRECONFIGURED_USA)} states active, "
          f"{len(PRECONFIGURED_DC)} extra jurisdiction inactive by default)")


if __name__ == "__main__":
    main()
