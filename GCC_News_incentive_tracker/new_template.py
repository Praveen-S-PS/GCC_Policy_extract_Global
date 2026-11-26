"""Convert the latest gcc_incentives_global_*.xlsx into the flat "new_template"
schema and save as output/new_template_GCC_policy_<date>.xlsx.

Also extracts sources/india/tamil_nadu.pdf directly so Tamil Nadu content is
included even if the main pipeline hasn't run since the PDF was added.

Run:  python3 new_template.py
"""

import difflib
import glob
import os
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from extractor.policy_pdf_reader import PolicyPDFReader
from extractor.policy_bucketer import PolicyBucketer
from feeds.sources_loader import SourcesLoader

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    requests = None
    BeautifulSoup = None


SOURCES_FILE = "input/sources/sources.xlsx"
INDIA_SOURCES_DIR = "sources/india"
GLOBAL_SOURCES_DIR = "sources/global"


INPUT_TEMPLATE = "input/new_template.xlsx"

# Split so the two folder families introduced by country_extractor.py stay
# separate: the wide/audit-trail workbook (old_template) that this module
# reads back in run(), and the flat signal/other workbook (new_template)
# that run() writes.
WIDE_SEARCH_DIR = "output/old_template/Global"
FLAT_OUTPUT_DIR = "output/new_template/Global"

# Backwards-compatible alias -- kept because it's still the right answer for
# "where do I write a flat workbook" when this module is run standalone.
OUTPUT_DIR = FLAT_OUTPUT_DIR

TAMIL_NADU_PDF = "sources/india/tamil_nadu.pdf"

TAMIL_NADU_VALIDITY = "Apr 2024 – Mar 2027"


HEADERS = [
    "section_name", "scope_type", "country_name", "state_name",
    "city_name", "tier_scope", "description",
    "validity_start_date", "validity_start_precision",
    "validity_end_date", "validity_end_precision",
    "gcc_relevance",
]


# Countries that formally have GCC-specific or near-GCC-specific policies.
# Rows from these get a +1 base score.
GCC_ANCHOR_COUNTRIES = {"India", "Mexico", "Poland"}


# Weighted signals of GCC/BSS relevance. Higher weight = stronger signal.
GCC_STRONG_SIGNALS = (
    "global capability centre", "global capability center",
    "capability centre", "capability center",
    "shared services", "shared service center", "shared service centre",
    "captive centre", "captive center", "captive unit",
    "gbs", "bss", "bpo", "kpo", "it/ites", "ites", "itbpm",
    "business services sector", "business services industry",
    "it services", "it services sector",
    "r&d center", "r&d centre", "research and development center",
    "research and development centre",
    "gcc unit", "gic",
    # 2026-09 review additions
    "global in-house centre", "global in-house center",
    "delivery centre", "delivery center",
    "global delivery centre", "global delivery center",
    "offshore development centre", "offshore development center", "odc",
    "engineering centre", "engineering center",
    "technology centre", "technology center",
    "competency centre", "competency center",
    "back office", "middle office", "regional headquarters",
)

# bss, gic and coe are short enough to also mean something unrelated (bss is
# a common telecom term, gic is a Singaporean sovereign fund, coe a
# center-of-excellence abbreviation used far outside this context). 2026-09
# review: require these three to be corroborated by another signal on the
# same page rather than counting alone -- word-boundary matching alone
# isn't enough since all three are also valid whole words in their other
# meanings.
GCC_AMBIGUOUS_SHORT_SIGNALS = {"bss", "gic", "coe"}

GCC_MEDIUM_SIGNALS = (
    "innovation hub", "innovation lab", "centre of excellence",
    "center of excellence", "coe",
    "ip box", "patent box", "r&d relief", "r&d grant", "r&d tax",
    "innovation grant", "digital services", "technology park",
    "it park", "software park",
)

GCC_WEAK_SIGNALS = (
    "innovation", "digital transformation", "outsourcing",
    "offshore", "nearshore", "multinational", "mnc",
    "shared services centre", "software", "call center", "call centre",
)


def _signal_matches(signals, text):
    return [kw for kw in signals if re.search(rf"\b{re.escape(kw)}\b", text)]


def compute_gcc_relevance(description, country_name):
    """Score 0..5+ for how strongly a row reads as GCC-applicable content.

    +1 if the jurisdiction is on the small list of countries with actual
    GCC-specific or near-GCC-specific policies.
    +2 per strong signal (business services sector, capability centre,
       shared services, ITES, BPO, R&D center)
    +1 per medium signal (IP box, R&D relief, innovation hub)
    +0.5 per weak signal (innovation, software, outsourcing)

    Matching is whole-word (not substring) so short tokens like "coe" don't
    fire inside unrelated words. bss/gic/coe additionally require a second
    signal elsewhere in the text -- see GCC_AMBIGUOUS_SHORT_SIGNALS.
    """

    if not description:
        return 0

    text = description.lower()

    score = 0.0

    if country_name in GCC_ANCHOR_COUNTRIES:
        score += 1

    strong_hits = _signal_matches(GCC_STRONG_SIGNALS, text)
    medium_hits = _signal_matches(GCC_MEDIUM_SIGNALS, text)
    weak_hits = _signal_matches(GCC_WEAK_SIGNALS, text)

    unambiguous_strong = [kw for kw in strong_hits if kw not in GCC_AMBIGUOUS_SHORT_SIGNALS]
    ambiguous_strong = [kw for kw in strong_hits if kw in GCC_AMBIGUOUS_SHORT_SIGNALS]
    unambiguous_medium = [kw for kw in medium_hits if kw not in GCC_AMBIGUOUS_SHORT_SIGNALS]
    ambiguous_medium = [kw for kw in medium_hits if kw in GCC_AMBIGUOUS_SHORT_SIGNALS]

    # "coe" can appear in both tiers (it's a corroborating hit for itself in
    # the other tier too, but two ambiguous hits of the SAME token in one
    # description shouldn't count as two independent signals).
    distinct_ambiguous = set(ambiguous_strong) | set(ambiguous_medium)

    has_corroboration = bool(
        unambiguous_strong or unambiguous_medium or weak_hits
        or len(distinct_ambiguous) > 1
    )

    if unambiguous_strong or (ambiguous_strong and has_corroboration):
        score += 2

    if unambiguous_medium or (ambiguous_medium and has_corroboration):
        score += 1

    if weak_hits:
        score += 0.5

    return int(round(score))


# Map incentive type names from the workbook to new_template section_name.
# Note: new_template capitalises "Non-Financial" (F upper), workbook uses lower.
SECTION_MAP = {
    "Financial Incentives": "Financial Incentives",
    "Infrastructure": "Infrastructure",
    "Talent": "Talent",
    "Regulatory, Ease-of-Doing-Business & Non-financial Support":
        "Regulatory, Ease-of-Doing-Business & Non-Financial Support",
    "Innovation / R&D / Patent": "Innovation / R&D / Patent",
    "Others": "Others",
}


# Map bucketer scope labels to (scope_type, tier_scope) for the new template.
# tier_scope is left blank for "Across ... (Common)" per user request.
SCOPE_MAP_INDIA = {
    "Across the state (Common)": ("State", ""),
    "Specific to Tier 1 cities": ("Tier", "Tier 1"),
    "Specific to Tier 2 & 3 cities": ("Tier", "Tier 2 & 3"),
}

SCOPE_MAP_GLOBAL = {
    "Across the Country (Common)": ("Country", ""),
    "Across the state (Common)": ("Country", ""),
    "Specific to Tier 1 cities": ("Tier", "Tier 1"),
    "Specific to Tier 2 & 3 cities": ("Tier", "Tier 2 & 3"),
}


# City registries: when a bullet mentions one of these city names we upgrade
# scope_type to "City" and fill city_name + tier_scope automatically.
INDIAN_STATE_CITIES = {
    "Andhra Pradesh": {
        "Tier 1": ["Visakhapatnam"],
        "Tier 2 & 3": ["Vijayawada", "Guntur", "Tirupati"],
    },
    "Gujarat": {
        "Tier 1": ["Ahmedabad", "Gandhinagar", "GIFT City", "Vadodara"],
        "Tier 2 & 3": ["Surat", "Rajkot", "Bhavnagar"],
    },
    "Karnataka": {
        "Tier 1": ["Bengaluru", "Bangalore"],
        "Tier 2 & 3": ["Mysuru", "Mysore", "Mangaluru", "Mangalore",
                       "Hubballi", "Belagavi", "Belgaum", "Shivamogga"],
    },
    "Kerala": {
        "Tier 1": ["Kochi", "Ernakulam"],
        "Tier 2 & 3": ["Thiruvananthapuram", "Kozhikode", "Thrissur"],
    },
    "Madhya Pradesh": {
        "Tier 1": ["Indore", "Bhopal"],
        "Tier 2 & 3": ["Gwalior", "Jabalpur", "Ujjain"],
    },
    "Maharashtra": {
        "Tier 1": ["Mumbai", "Pune", "Navi Mumbai"],
        "Tier 2 & 3": ["Nagpur", "Nashik", "Aurangabad", "Kolhapur",
                       "Solapur", "Amravati"],
    },
    "Rajasthan": {
        "Tier 1": ["Jaipur"],
        "Tier 2 & 3": ["Udaipur", "Jodhpur", "Kota", "Ajmer"],
    },
    "Tamil Nadu": {
        "Tier 1": ["Chennai"],
        "Tier 2 & 3": ["Coimbatore", "Madurai", "Tiruchirappalli", "Trichy",
                       "Salem", "Tirunelveli", "Hosur", "Erode", "Vellore",
                       "Tuticorin", "Thoothukudi"],
    },
    "Telangana": {
        "Tier 1": ["Hyderabad"],
        "Tier 2 & 3": ["Warangal", "Nizamabad", "Karimnagar"],
    },
    "Uttar Pradesh": {
        "Tier 1": ["Noida", "Ghaziabad", "Gautam Buddha Nagar", "Lucknow"],
        "Tier 2 & 3": ["Kanpur", "Varanasi", "Prayagraj", "Agra",
                       "Meerut", "Bareilly"],
    },
    "West Bengal": {
        "Tier 1": ["Kolkata"],
        "Tier 2 & 3": ["Siliguri", "Durgapur", "Asansol"],
    },
}


COUNTRY_CITIES = {
    "United Kingdom": {
        "Tier 1": ["London"],
        "Tier 2 & 3": ["Manchester", "Birmingham", "Edinburgh", "Leeds",
                       "Bristol", "Glasgow", "Belfast", "Cardiff"],
    },
    "Germany": {
        "Tier 1": ["Berlin", "Frankfurt", "Munich", "München"],
        "Tier 2 & 3": ["Hamburg", "Cologne", "Köln", "Stuttgart",
                       "Düsseldorf", "Leipzig"],
    },
    "France": {
        "Tier 1": ["Paris"],
        "Tier 2 & 3": ["Lyon", "Marseille", "Toulouse", "Nice", "Nantes",
                       "Bordeaux"],
    },
    "Japan": {
        "Tier 1": ["Tokyo", "Osaka"],
        "Tier 2 & 3": ["Yokohama", "Nagoya", "Fukuoka", "Kobe", "Kyoto",
                       "Sapporo"],
    },
    "Philippines": {
        "Tier 1": ["Manila", "Metro Manila", "Makati", "Taguig", "BGC",
                   "Bonifacio Global City", "Cebu City", "Cebu"],
        "Tier 2 & 3": ["Davao", "Iloilo", "Bacolod", "Cagayan de Oro",
                       "Clark", "Bacoor"],
    },
    "Malaysia": {
        "Tier 1": ["Kuala Lumpur", "Cyberjaya", "Putrajaya"],
        "Tier 2 & 3": ["Penang", "Georgetown", "Iskandar Puteri",
                       "Johor Bahru", "Kuching"],
    },
    "Mexico": {
        "Tier 1": ["Mexico City", "CDMX", "Monterrey", "Guadalajara"],
        "Tier 2 & 3": ["Puebla", "Querétaro", "Queretaro", "Tijuana",
                       "Aguascalientes"],
    },
    "Poland": {
        "Tier 1": ["Warsaw", "Warszawa", "Kraków", "Krakow", "Wrocław",
                   "Wroclaw", "Gdańsk", "Gdansk"],
        "Tier 2 & 3": ["Katowice", "Poznań", "Poznan", "Łódź", "Lodz",
                       "Lublin", "Rzeszów", "Rzeszow"],
    },
    "Vietnam": {
        "Tier 1": ["Ho Chi Minh City", "Ho Chi Minh", "Hanoi", "Ha Noi"],
        "Tier 2 & 3": ["Da Nang", "Haiphong", "Hai Phong", "Can Tho"],
    },
    "Costa Rica": {
        "Tier 1": ["San José", "San Jose", "Heredia"],
        "Tier 2 & 3": ["Alajuela", "Cartago"],
    },
    "Colombia": {
        "Tier 1": ["Bogotá", "Bogota", "Medellín", "Medellin"],
        "Tier 2 & 3": ["Cali", "Barranquilla", "Cartagena"],
    },
    "Brazil": {
        "Tier 1": ["São Paulo", "Sao Paulo", "Rio de Janeiro", "Brasília",
                   "Brasilia"],
        "Tier 2 & 3": ["Curitiba", "Porto Alegre", "Recife", "Fortaleza",
                       "Belo Horizonte"],
    },
    "Argentina": {
        "Tier 1": ["Buenos Aires", "CABA"],
        "Tier 2 & 3": ["Córdoba", "Cordoba", "Rosario", "Mendoza"],
    },
    "Chile": {
        "Tier 1": ["Santiago"],
        "Tier 2 & 3": ["Valparaíso", "Valparaiso", "Concepción", "Concepcion"],
    },
    "Egypt": {
        "Tier 1": ["Cairo", "Alexandria", "New Administrative Capital"],
        "Tier 2 & 3": ["Giza", "Ismailia"],
    },
    "South Africa": {
        "Tier 1": ["Johannesburg", "Cape Town", "Pretoria", "Sandton"],
        "Tier 2 & 3": ["Durban", "Port Elizabeth", "Gqeberha", "Bloemfontein"],
    },
    "UAE": {
        "Tier 1": ["Dubai", "Abu Dhabi"],
        "Tier 2 & 3": ["Sharjah", "Ras Al Khaimah", "Ajman"],
    },
    "Saudi Arabia": {
        "Tier 1": ["Riyadh", "Jeddah"],
        "Tier 2 & 3": ["Dammam", "NEOM", "Al Khobar"],
    },
    "Kenya": {
        "Tier 1": ["Nairobi"],
        "Tier 2 & 3": ["Mombasa", "Kisumu", "Konza"],
    },
    "Morocco": {
        "Tier 1": ["Casablanca", "Rabat"],
        "Tier 2 & 3": ["Tangier", "Marrakech", "Fez"],
    },
    "Indonesia": {
        "Tier 1": ["Jakarta"],
        "Tier 2 & 3": ["Surabaya", "Bandung", "Medan", "Batam"],
    },
    "Thailand": {
        "Tier 1": ["Bangkok"],
        "Tier 2 & 3": ["Chiang Mai", "Phuket", "Chonburi"],
    },
    "Singapore": {
        "Tier 1": ["Singapore"],
        "Tier 2 & 3": [],
    },
    "Bangladesh": {
        "Tier 1": ["Dhaka"],
        "Tier 2 & 3": ["Chittagong", "Chattogram"],
    },
    "Turkey": {
        "Tier 1": ["Istanbul", "Ankara"],
        "Tier 2 & 3": ["Izmir", "Bursa"],
    },
}


MONTHS = {
    # English
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
    # Spanish
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5,
    "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9,
    "octubre": 10, "noviembre": 11, "diciembre": 12,
    # Portuguese
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3,
    "maio": 5, "junho": 6, "julho": 7,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12,
    # French
    "janvier": 1, "février": 2, "fevrier": 2, "mars": 3,
    "avril": 4, "mai": 5, "juin": 6, "juillet": 7,
    "août": 8, "aout": 8, "septembre": 9,
    "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12,
    # Polish (nominative + genitive after od/do)
    "styczeń": 1, "stycznia": 1,
    "luty": 2, "lutego": 2,
    "marzec": 3, "marca": 3,
    "kwiecień": 4, "kwietnia": 4,
    "maj": 5, "maja": 5,
    "czerwiec": 6, "czerwca": 6,
    "lipiec": 7, "lipca": 7,
    "sierpień": 8, "sierpnia": 8,
    "wrzesień": 9, "września": 9,
    "październik": 10, "października": 10,
    "listopad": 11, "listopada": 11,
    "grudzień": 12, "grudnia": 12,
}


LAST_DAY_OF_MONTH = {
    1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30,
    7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31,
}


DATE_SEP = re.compile(r"\s*[-–—]\s*")

BULLET_SPLIT = "\n\n• "


# Patterns that mine dates out of cell text itself. Tried in this order.
# MONTH_NAME is built dynamically from the multilingual MONTHS dict so
# adding a new month name in MONTHS above automatically extends the regex.
MONTH_NAME = "(?:" + "|".join(
    re.escape(k) for k in sorted(MONTHS.keys(), key=len, reverse=True)
) + ")"

# Optional explicit day number around a month name, captured (not discarded)
# so an exact day like "11 February" or "March 31," can override the
# month-precision default (day=1 for a start, last-day-of-month for an end).
_LEAD_DAY = r"(?:(\d{1,2})(?:st|nd|rd|th|er|e)?\s+(?:de\s+|do\s+|d[ea]\s+|of\s+)?)?"
_TRAIL_DAY = r"(?:\s+(\d{1,2})(?:st|nd|rd|th)?,?)?"


# Language-agnostic "from" and "until" trigger words used before a month/year.
FROM_TRIGGERS = (
    r"(?:with effect from|effective from|w\.?e\.?f\.?|starting from|"
    r"commencing (?:on|from)|"
    r"in force (?:from|w\.?e\.?f\.?|since|with effect from)|"
    r"valid from|has been in effect since|"
    # US legislative phrasing: "enacted in 2016", "signed into law in
    # July 2025" -- common on US state incentive-program pages, which
    # describe a program's creation rather than a formal "effective from".
    r"enacted (?:in|on)|was enacted (?:in|on)|signed into law (?:in|on)|"
    # Spanish
    r"vigente desde|a partir de|en vigor desde|"
    # Portuguese
    r"em vigor (?:a partir de|desde)|válido (?:a partir de|desde)|"
    r"entrada em vigor em|"
    # French
    r"en vigueur (?:du|à compter du|depuis le)|à compter du|"
    r"applicable à (?:partir|compter) du|"
    # Polish
    r"obowiązuje od|w mocy od|z dniem)"
)

UNTIL_TRIGGERS = (
    # NOTE: a bare "by" is deliberately excluded -- government documents
    # constantly use "by <year>" for aspirational targets/vision statements
    # (e.g. "reach $110 billion by 2030", "developed nation by 2047"),
    # which is a false-positive magnet for a policy end date.
    r"(?:valid until|valid till|(?:expires?|sunsets?) (?:in|on|by)|"
    r"set to expire (?:in|on)?|until|till|"
    r"applicable until|no later than|by end of|"
    # Spanish
    r"vigente hasta|válido hasta|hasta el|hasta|"
    # Portuguese
    r"válido até|em vigor até|até (?:o dia )?|"
    # French
    r"jusqu[’']au|jusqu[’']à|valable jusqu[’']au|en vigueur jusqu[’']au|"
    # Polish
    r"do dnia|obowiązuje do|w mocy do)"
)

# Range connectors in multiple languages
RANGE_CONNECTOR = r"(?:to|through|until|hasta|até|au|à|do|-|–|—)"


# 1. Month YYYY <sep> Month YYYY. Both sides allow an optional day number
# either before the month (e.g. "1 janvier 2024 au 31 décembre 2027") or
# after it (e.g. "March 31, 2029"); whichever is present is captured so the
# exact day survives instead of being discarded.
RE_MONTH_YEAR_RANGE = re.compile(
    rf"{_LEAD_DAY}({MONTH_NAME}){_TRAIL_DAY}\.?\s*,?\s+(?:de\s+|do\s+)?(\d{{4}})"
    rf"\s*{RANGE_CONNECTOR}\s*"
    rf"{_LEAD_DAY}({MONTH_NAME}){_TRAIL_DAY}\.?\s*,?\s+(?:de\s+|do\s+)?(\d{{4}})",
    re.IGNORECASE,
)

# 2. YYYY <sep> YYYY (must not be adjacent to "of" or currency)
RE_YEAR_RANGE = re.compile(
    rf"(?<![A-Za-z$€₹])(20\d{{2}})\s*{RANGE_CONNECTOR}\s*(20\d{{2}})\b"
)

# 3. Effective from Month YYYY (multilingual). Captures an explicit day
# before or after the month if the source states one.
RE_FROM_MONTH_YEAR = re.compile(
    rf"\b{FROM_TRIGGERS}\s+"
    rf"(?:the\s+|le\s+|el\s+|o\s+)?"
    rf"{_LEAD_DAY}"
    rf"({MONTH_NAME})"
    rf"{_TRAIL_DAY}"
    rf"\.?\s*,?\s+(?:de\s+|do\s+)?(\d{{4}})",
    re.IGNORECASE,
)

# 4. Effective from YYYY / a partir de YYYY (year only)
RE_FROM_YEAR = re.compile(
    rf"\b{FROM_TRIGGERS}\s+(?:the\s+|el\s+|o\s+|le\s+)?(20\d{{2}})\b",
    re.IGNORECASE,
)

# 5. Valid until / expires in Month YYYY (multilingual). Captures an
# explicit day before or after the month (e.g. "till 31st March 2030" or
# "valid until March 31, 2029") if the source states one.
RE_UNTIL_MONTH_YEAR = re.compile(
    rf"\b{UNTIL_TRIGGERS}\s+"
    rf"(?:the\s+|le\s+|el\s+|o\s+)?"
    rf"{_LEAD_DAY}"
    rf"({MONTH_NAME})"
    rf"{_TRAIL_DAY}"
    rf"\.?\s*,?\s+(?:de\s+|do\s+)?(\d{{4}})",
    re.IGNORECASE,
)

# 6. Valid until YYYY / vigente hasta YYYY (year only)
RE_UNTIL_YEAR = re.compile(
    rf"\b{UNTIL_TRIGGERS}\s+(?:the\s+|el\s+|o\s+|le\s+)?(20\d{{2}})\b",
    re.IGNORECASE,
)

# 7. For N years starting <year>
RE_DURATION_FROM_YEAR = re.compile(
    r"\bfor\s+(\d{1,2})\s+years?\s+(?:starting\s+)?(?:from|since)\s+(20\d{2})\b",
    re.IGNORECASE,
)

# 8. Policy 2024-30, 2025-2030 shortform, "Policy 2025-30" etc.
# Requires a literal hyphen directly joining the two numbers (only spaces/
# tabs allowed around it, no newline) so this can't stretch across a page
# break into an unrelated section number, e.g. "...GCC Policy 2024\n12.
# Definitions" must NOT be read as a "2024-2012" policy range.
RE_POLICY_SHORT = re.compile(
    r"\bpolicy\s+(20\d{2})[ \t]*-[ \t]*(\d{2,4})\b",
    re.IGNORECASE,
)

# 9. Labeled dates: "Effective date: 15/03/2024", "Validity: 2024-01-01 to 2027-12-31"
RE_LABELED_START = re.compile(
    r"\b(?:effective\s+date|validity\s+start|start\s+date|notified\s+on|"
    r"in\s+force\s+from|date\s+of\s+notification|commencement\s+date)"
    r"\s*[:\-]?\s*"
    r"(\d{1,2})[\s/\-.](\d{1,2})[\s/\-.](\d{4})",
    re.IGNORECASE,
)

RE_LABELED_END = re.compile(
    r"\b(?:expiry\s+date|expiration\s+date|end\s+date|validity\s+end|"
    r"sunset\s+(?:date|clause)?|valid\s+till|valid\s+until|until\s+date)"
    r"\s*[:\-]?\s*"
    r"(\d{1,2})[\s/\-.](\d{1,2})[\s/\-.](\d{4})",
    re.IGNORECASE,
)

# 9b. Bare "Dated: DD/MM/YYYY" letterhead line, common in Indian government
# resolutions/orders. Requires the colon so it doesn't match a mid-sentence
# reference to an unrelated document (e.g. "...notification issued by the
# Labour Department dated 24.06.2016" has no colon and is correctly ignored).
RE_DATED_COLON = re.compile(
    r"\bdated\s*:\s*(\d{1,2})[\s/\-.](\d{1,2})[\s/\-.](\d{4})",
    re.IGNORECASE,
)

# 10. ISO-style date range: 2024-01-01 to 2027-12-31
RE_ISO_DATE_RANGE = re.compile(
    r"\b(20\d{2})-(\d{1,2})-(\d{1,2})\s*(?:to|until|through|[-–—])\s*"
    r"(20\d{2})-(\d{1,2})-(\d{1,2})\b"
)

# 11. Duration phrase: "for a period of N years from Month YYYY"
RE_PERIOD_FROM = re.compile(
    rf"for\s+a\s+period\s+of\s+(\d{{1,2}})\s+years?\s+"
    rf"(?:starting|commencing)?\s*(?:from|since)\s+"
    rf"({MONTH_NAME})\.?\s+(?:de\s+|do\s+)?(\d{{4}})",
    re.IGNORECASE,
)


# Signals that a nearby date range is about policy validity rather than
# historical statistics. Used to guard the year-range regex from picking up
# FDI stats like "2016 through 2023" that aren't the policy's own validity.
POLICY_CONTEXT_RE = re.compile(
    r"\b(policy|scheme|law|act|regulation|decree|order|resolution|"
    r"valid|validity|effective|effect|in force|in effect|"
    r"expire|expiry|expires|expiring|"
    r"duration|period|implementation|operative|"
    r"applicable|applicability|"
    r"gazette|notification|notified|"
    r"from date of|w\.?e\.?f\.?|until 31|through 31|"
    # Spanish
    r"vigente|vigencia|decreto|ley|reglamento|"
    # Portuguese
    r"vigente|vigência|em vigor|decreto|lei|"
    # French
    r"en vigueur|décret|loi|règlement|arrêté|"
    # Polish
    r"obowiązuje|ustawa|rozporządzenie|dekret)\b",
    re.IGNORECASE,
)

# Signals that a nearby date range is about statistics or historical data.
# When these appear near a match we refuse to accept it as a policy validity.
STATS_CONTEXT_RE = re.compile(
    r"\b(fdi|foreign direct investment|attracted|reached|grew|growth|"
    r"totalled|totaled|amounted|averaged|invested|received|rose|"
    r"gdp|revenue|exports|imports|employment (?:reached|grew|rose)|"
    r"between .{0,20}? and)\b",
    re.IGNORECASE,
)


# Today's date used to filter expired policies (only when we have a
# confident end date; ongoing / blank ends are always kept).
TODAY = date.today()


HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")


def find_latest_workbook():

    files = [
        f for f in glob.glob(os.path.join(WIDE_SEARCH_DIR, "gcc_incentives_global_*.xlsx"))
        if not os.path.basename(f).startswith("~$")
        and not os.path.basename(f).startswith("new_template")
    ]

    if not files:
        return None

    # Prefer paraphrased variants when both are present
    para = [f for f in files if "paraphrased" in os.path.basename(f)]

    picks = para or files

    picks.sort(key=os.path.getmtime, reverse=True)

    return picks[0]


def parse_bullets(cell_value):

    if cell_value is None:
        return []

    text = str(cell_value).strip()

    if not text:
        return []

    if text.startswith("• "):
        text = text[2:]

    parts = text.split(BULLET_SPLIT)

    return [p.strip() for p in parts if p.strip() and len(p.strip()) > 20]


def parse_validity(s):
    """Return (start_date, start_prec, end_date, end_prec)."""

    if not s or not str(s).strip():
        return ("", "", "", "")

    s = str(s).strip()

    parts = DATE_SEP.split(s, maxsplit=1)

    if len(parts) == 1:
        d, p = _parse_single_date(parts[0])
        return (d, p, "", "")

    start_d, start_p = _parse_single_date(parts[0])
    end_d, end_p = _parse_single_date(parts[1], is_end=True)

    return (start_d, start_p, end_d, end_p)


def _parse_single_date(s, is_end=False):
    """Parse one side of a "Mon YYYY - Mon YYYY" style Validity-column entry.

    When only a month/year is given (no explicit day), the end side
    defaults to the last day of that month rather than the 1st -- "valid
    ... - Mar 2030" reads as remaining in force through March, not expiring
    on its first day. The start side keeps defaulting to day 1."""

    s = s.strip()

    low = s.lower()

    if low in ("ongoing", "onwards", "present", "current"):
        return (low, "")

    m = re.match(r"([A-Za-z]{3,})\.?\s*(\d{4})", s)

    if m:

        month_key = m.group(1).lower().rstrip(".")

        month = MONTHS.get(month_key) or MONTHS.get(month_key[:3], 1)

        return (_mk_month_date(m.group(2), month, end_of_month=is_end), "month")

    m = re.match(r"(\d{4})", s)

    if m:
        year = m.group(1)
        return ((f"{year}-12-31" if is_end else f"{year}-01-01"), "year")

    return (s, "")


def _month_number(name):

    key = name.lower().rstrip(".")

    return MONTHS.get(key) or MONTHS.get(key[:3], 1)


def _mk_month_date(year, month, day=None, end_of_month=False):

    if day:
        d = int(day)
    elif end_of_month:
        d = LAST_DAY_OF_MONTH.get(month, 28)
    else:
        d = 1

    return f"{year}-{month:02d}-{d:02d}"


def extract_validity_from_text(text):
    """Scan a bullet's text for a validity range. Returns
    (start_date, start_precision, end_date, end_precision) with empty strings
    when nothing plausible is found.

    Patterns are tried in order of confidence, and each one only fills in
    whichever of start/end is still blank rather than stopping at the first
    match. Explicit labeled dates ("Dated: 11/02/2025", "effective from
    ...", "valid until ...") are tried before a bare "YYYY-YYYY" range,
    because a bare range can latch onto an unrelated policy mentioned in
    passing nearby (e.g. "... Biotechnology Policy (2020-2025) ...").
    """

    if not text:
        return ("", "", "", "")

    start, s_prec, end, e_prec = "", "", "", ""

    # 1. Month YYYY - Month YYYY in one clause, optionally with an exact
    # day on either side. An explicit paired range is the highest-confidence
    # single match, so it returns immediately.
    m = RE_MONTH_YEAR_RANGE.search(text)

    if m:
        s_lead, sm, s_trail, sy, e_lead, em, e_trail, ey = m.groups()
        sm_n = _month_number(sm)
        em_n = _month_number(em)
        s_day = s_lead or s_trail
        e_day = e_lead or e_trail
        return (
            _mk_month_date(sy, sm_n, day=s_day), "day" if s_day else "month",
            _mk_month_date(ey, em_n, day=e_day, end_of_month=True),
            "day" if e_day else "month",
        )

    # 2. Explicit labeled / trigger-word dates: "Dated: ...", "effective
    # date/from ...", "valid until/till ...", "notified on ...". Tried
    # before the bare year-range below since these are unambiguous about
    # which policy they describe.
    start_pair = _find_start(text)
    end_pair = _find_end(text)

    if start_pair:
        start, s_prec = start_pair
    if end_pair:
        end, e_prec = end_pair

    if start and end:
        return (start, s_prec, end, e_prec)

    # 3. Bare YYYY - YYYY (only accepted when policy language is nearby and
    # statistics language is NOT nearby, otherwise it's likely a historical
    # dataset range, or a different policy's own year range mentioned in
    # passing). Only fills whichever side is still missing.
    m = RE_YEAR_RANGE.search(text)

    if m and _is_policy_context(text, m):
        sy, ey = m.group(1), m.group(2)
        start = start or f"{sy}-01-01"
        s_prec = s_prec or "year"
        end = end or f"{ey}-12-31"
        e_prec = e_prec or "year"

    if start and end:
        return (start, s_prec, end, e_prec)

    # 3b. Bare "from/until <year>" with no month attached -- the weakest,
    # lowest-confidence signal, tried only now. A bare trigger + year shows
    # up constantly in non-validity sentences (market-outlook captions,
    # infographic timelines like "Till 2010 2011-2015 ... 2024 Onwards"),
    # so it must not outrank the policy-context-guarded range above.
    if not start:
        p = _find_start_year_only(text)
        if p:
            start, s_prec = p
    if not end:
        p = _find_end_year_only(text)
        if p:
            end, e_prec = p

    if start and end:
        return (start, s_prec, end, e_prec)

    # 4. Duration: "for N years starting/since YYYY"
    m = RE_DURATION_FROM_YEAR.search(text)
    if m:
        years = int(m.group(1))
        sy = int(m.group(2))
        start = start or f"{sy}-01-01"
        s_prec = s_prec or "year"
        end = end or f"{sy + years}-12-31"
        e_prec = e_prec or "year"

    if start and end:
        return (start, s_prec, end, e_prec)

    # 5. Policy short form like "Policy 2025-30"
    m = RE_POLICY_SHORT.search(text)
    if m:
        sy = m.group(1)
        end_raw = m.group(2)
        ey = (
            str(int(sy[:2]) * 100 + int(end_raw))
            if len(end_raw) == 2 else end_raw
        )
        start = start or f"{sy}-01-01"
        s_prec = s_prec or "year"
        end = end or f"{ey}-12-31"
        e_prec = e_prec or "year"

    if start and end:
        return (start, s_prec, end, e_prec)

    # 6. ISO-style range: 2024-01-01 to 2027-12-31
    m = RE_ISO_DATE_RANGE.search(text)
    if m:
        isy, ism, isd, iey, iem, ied = m.groups()
        start = start or f"{isy}-{int(ism):02d}-{int(isd):02d}"
        s_prec = s_prec or "day"
        end = end or f"{iey}-{int(iem):02d}-{int(ied):02d}"
        e_prec = e_prec or "day"

    if start and end:
        return (start, s_prec, end, e_prec)

    # 7. Period expression: "for a period of 5 years from January 2024"
    m = RE_PERIOD_FROM.search(text)
    if m:
        years = int(m.group(1))
        sm_n = _month_number(m.group(2))
        sy = int(m.group(3))
        start = start or _mk_month_date(str(sy), sm_n)
        s_prec = s_prec or "month"
        end = end or _mk_month_date(str(sy + years), sm_n, end_of_month=True)
        e_prec = e_prec or "month"

    return (start, s_prec, end, e_prec)


def _find_start(text):
    """High-confidence start-date patterns only: an explicit numeric label
    ("Dated: ...") or an explicit month name ("effective from March 2025").
    Deliberately excludes the bare "from <year>" trigger with no month --
    see _find_start_year_only, which is tried later at lower priority."""

    m = RE_LABELED_START.search(text) or RE_DATED_COLON.search(text)
    if m:
        d, mo, y = m.groups()
        return (f"{y}-{int(mo):02d}-{int(d):02d}", "day")

    m = RE_FROM_MONTH_YEAR.search(text)
    if m:
        lead_day, month_name, trail_day, year = m.groups()
        day = lead_day or trail_day
        month_n = _month_number(month_name)
        return (
            _mk_month_date(year, month_n, day=day),
            "day" if day else "month",
        )

    return None


def _find_end(text):
    """High-confidence end-date patterns only: an explicit numeric label or
    an explicit month name. See _find_start's note -- the bare-year trigger
    is intentionally excluded here too."""

    m = RE_LABELED_END.search(text)
    if m:
        d, mo, y = m.groups()
        return (f"{y}-{int(mo):02d}-{int(d):02d}", "day")

    m = RE_UNTIL_MONTH_YEAR.search(text)
    if m:
        lead_day, month_name, trail_day, year = m.groups()
        day = lead_day or trail_day
        month_n = _month_number(month_name)
        return (
            _mk_month_date(year, month_n, day=day, end_of_month=True),
            "day" if day else "month",
        )

    return None


def _find_start_year_only(text):
    """Lowest-confidence start pattern: a bare "from/since <year>" with no
    month attached. Tried only after the policy-context-guarded YYYY-YYYY
    range (case 3 in extract_validity_from_text), because a bare trigger +
    year is common in non-validity sentences too (market-outlook captions,
    infographic timelines, etc.)."""

    m = _first_non_stats_match(RE_FROM_YEAR, text)
    if m:
        return (f"{m.group(1)}-01-01", "year")

    return None


def _find_end_year_only(text):
    """Lowest-confidence end pattern -- see _find_start_year_only."""

    m = _first_non_stats_match(RE_UNTIL_YEAR, text)
    if m:
        return (f"{m.group(1)}-12-31", "year")

    return None


def _is_policy_context(text, match):
    """Check text around a regex match for policy language and reject
    statistics language."""

    start, end = match.span()

    window_start = max(0, start - 100)
    window_end = min(len(text), end + 100)
    window = text[window_start:window_end]

    if STATS_CONTEXT_RE.search(window):
        return False

    return bool(POLICY_CONTEXT_RE.search(window))


_YEAR_TOKEN_RE = re.compile(r"\b(?:19|20)\d{2}\b")


def _looks_like_timeline_caption(text, match):
    """Reject a bare-year match sitting inside a chart/timeline caption --
    identifiable by several standalone year tokens packed close together
    with little connecting prose (e.g. a GCC-maturity-curve axis label
    like "Till 2010 2011-2015 2015-2024 2024 Onwards"). That is not a
    validity sentence, even though it satisfies a bare "till <year>"
    trigger."""

    window_start = max(0, match.start() - 40)
    window_end = min(len(text), match.end() + 40)

    return len(_YEAR_TOKEN_RE.findall(text[window_start:window_end])) >= 3


def _first_non_stats_match(pattern, text):
    """Like pattern.search(text), but skips over any match that sits next to
    statistics/target language (e.g. "reach $110 billion by 2030", "developed
    nation by 2047") or inside a chart/timeline caption, and returns the
    first match that doesn't. These bare year-only patterns have no month
    to anchor them, so they're the most prone to misreading an unrelated
    year as a policy date."""

    for m in pattern.finditer(text):

        window_start = max(0, m.start() - 100)
        window_end = min(len(text), m.end() + 100)

        if STATS_CONTEXT_RE.search(text[window_start:window_end]):
            continue

        if _looks_like_timeline_caption(text, m):
            continue

        return m

    return None


def is_expired(end_date_str, end_precision):
    """True when a confidently-known end date is before today."""

    if not end_date_str or end_date_str.lower() in ("ongoing", "onwards",
                                                     "present", "current"):
        return False

    try:
        end_dt = datetime.strptime(end_date_str, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return False

    return end_dt < TODAY


def scan_source_for_validity(path_or_url, is_url=False):
    """Read a PDF file or fetch an HTML URL and run the date extractor over
    the raw text.

    Reads the WHOLE document before parsing (rather than growing the text
    page by page and parsing after each page). An incremental accumulate-
    then-parse loop was tried and dropped: a match's surrounding context
    window can look different depending on how much of the document has
    been read so far, so the same sentence could be accepted on one partial
    prefix and correctly rejected once the rest of the page was in view --
    i.e. the result depended on where the scan happened to stop, not on
    the document's actual content. Parsing the complete text once removes
    that whole class of bug.

    Returns (start, s_prec, end, e_prec) if any validity signal is found,
    otherwise ('','','','').
    """

    try:

        if is_url:

            if not requests or not BeautifulSoup:
                return ("", "", "", "")

            resp = requests.get(
                path_or_url, timeout=30,
                headers={"User-Agent": "Mozilla/5.0"},
            )

            if resp.status_code >= 400:
                return ("", "", "", "")

            soup = BeautifulSoup(resp.text, "lxml")

            for tag in soup(("script", "style", "noscript", "nav",
                             "header", "footer", "aside", "form", "iframe")):
                tag.decompose()

            text = soup.get_text(" ", strip=True)

            return extract_validity_from_text(text)

        else:

            if not pdfplumber or not os.path.exists(path_or_url):
                return ("", "", "", "")

            pages = []

            with pdfplumber.open(path_or_url) as pdf:

                for page in pdf.pages:

                    t = page.extract_text() or ""

                    if t.strip():
                        pages.append(t)

            return extract_validity_from_text("\n".join(pages))

    except Exception:
        return ("", "", "", "")


def build_source_dates_index(srcs):
    """Return dict[(region, jurisdiction)] → (start, s_prec, end, e_prec).

    Reads each configured source PDF and HTML URL in parallel and scans
    for a validity period expressed in the source itself.
    """

    dates_by_jurisdiction = {}
    lock = threading.Lock()

    def _remember(key, dates):
        # Only overwrite if the new dates have BOTH start and end
        with lock:
            if not dates_by_jurisdiction.get(key) and (dates[0] or dates[2]):
                dates_by_jurisdiction[key] = dates

    workers = int(os.environ.get("SCAN_WORKERS", "10"))

    print(f"  ⓘ scanning source files for validity dates ({workers} workers)")

    # Assemble every scan task first, then run them concurrently.
    tasks = []  # each: (region, jurisdiction, path_or_url, is_url)

    for state, pdfs in srcs["india_pdfs"].items():
        for filename, _, _ in pdfs:
            tasks.append(("India", state,
                          os.path.join(INDIA_SOURCES_DIR, filename), False))

    for state, urls in srcs["india_html"].items():
        for url in urls:
            if url:
                tasks.append(("India", state, url, True))

    for country, pdfs in srcs["country_pdfs"].items():
        country_dir = country.lower().replace(" ", "_")
        for filename, _, _ in pdfs:
            tasks.append(("Global", country,
                          os.path.join(GLOBAL_SOURCES_DIR,
                                       country_dir, filename), False))

    for country, urls in srcs["country_html"].items():
        for url in urls[:2]:  # cap HTML scans per country
            if url:
                tasks.append(("Global", country, url, True))

    def _run(task):
        region, jurisdiction, target, is_url = task
        dates = scan_source_for_validity(target, is_url=is_url)
        _remember((region, jurisdiction), dates)

    with ThreadPoolExecutor(max_workers=workers) as ex:
        list(ex.map(_run, tasks))

    print(f"  ⓘ source date scan found validity for "
          f"{len(dates_by_jurisdiction)} jurisdictions")

    return dates_by_jurisdiction


def detect_cities(text, cities_by_tier):
    """Find every city from `cities_by_tier` that appears in `text`.

    Returns list of (city_name, tier_scope) tuples. Empty list if no cities
    match or the registry is missing.
    """

    if not cities_by_tier:
        return []

    text_low = text.lower()

    matches = []

    for tier, cities in cities_by_tier.items():

        for city in cities:

            if re.search(rf"\b{re.escape(city.lower())}\b", text_low):

                matches.append((city, tier))

    # Deduplicate by (city, tier)
    seen = set()
    unique = []
    for entry in matches:
        if entry not in seen:
            seen.add(entry)
            unique.append(entry)

    return unique


def _extract_incentive_type_map(ws, incentive_col=1):
    """Walk the sheet from row 3 and track which incentive type applies to each
    row, handling merged cells that only carry the label on their first row.
    """

    current = ""

    by_row = {}

    for r in range(3, ws.max_row + 1):

        v = ws.cell(r, incentive_col).value

        if v and str(v).strip():
            current = str(v).strip()

        by_row[r] = current

    return by_row


def _apply_source_fallback(text_dates, source_dates, validity_dates):
    """Merge validity info in priority order, then default an empty end
    date to 'ongoing'.

    Priority: 1) dates found in the bullet's own text (most specific),
    2) dates found by re-scanning the full source PDF/HTML directly,
    3) the manually-entered Validity column in sources.xlsx.

    The Validity column is human-entered shorthand (e.g. "Feb 2025 - Mar
    2030") and is usually only month precision, or can go stale if the
    policy's actual notification date differs from what was known when the
    row was added. The full-document scan re-reads the primary source
    itself and can recover exact day-level dates (e.g. "Dated: 11/02/2025"),
    so it is preferred over the manual column whenever it finds something.

    Rationale for the "ongoing" default: every row we emit is backed by an
    active government or agency source. Blank end dates create ambiguity;
    'ongoing' communicates that the policy is currently in effect (unless a
    specific end has been found)."""

    t_s, t_sp, t_e, t_ep = text_dates or ("", "", "", "")
    src_s, src_sp, src_e, src_ep = source_dates or ("", "", "", "")
    col_s, col_sp, col_e, col_ep = validity_dates or ("", "", "", "")

    start = t_s or src_s or col_s
    s_prec = t_sp or src_sp or col_sp
    end = t_e or src_e or col_e
    e_prec = t_ep or src_ep or col_ep

    # Default any unknown end date to "ongoing" — the row represents a
    # currently-published policy source, so if no end is documented we
    # assume the policy is still in force.
    if not end:
        end = "ongoing"
        e_prec = ""

    return (start, s_prec, end, e_prec)


def _lcs_masks(text):
    """Per character bitmask of the positions it occurs at in text."""

    masks = {}

    for i, ch in enumerate(text):
        masks[ch] = masks.get(ch, 0) | (1 << i)

    return masks


def _lcs_length(masks, n, other):
    """Longest common subsequence length of a text (given as its _lcs_masks
    and length n) and `other`, by the bit parallel algorithm (Hyyro 2004):
    one big integer operation per character of `other`, so a pair of 3000
    character strings takes about a millisecond."""

    all_ones = (1 << n) - 1
    v = all_ones

    for ch in other:

        m = masks.get(ch)

        if m is None:
            continue

        u = v & m
        v = ((v + u) | (v - u)) & all_ones

    return n - bin(v).count("1")


def _fuzzy_dedupe_rows(rows, threshold=0.85):
    """Drop rows whose description is a near-duplicate of an earlier row's
    within the same (country, state, section, location) group.

    The exact-prefix dedup applied right after bucketing (country_extractor.
    py's _dedupe_items) catches byte-identical bullets, but not the same
    underlying source paragraph extracted twice (e.g. reached via both a PDF
    and a linked HTML page) and then paraphrased into two slightly different
    wordings -- paraphrasing an LLM call isn't deterministic, so two passes
    over near-identical input can diverge within the first 60 characters.
    Confirmed on a 16-country Global subset run: Spain, Thailand, Mauritius,
    Taiwan and South Korea each had 2-4 rows that were clearly the same fact
    reworded. Grouped by (country, state, section) to keep the O(n^2)
    comparison cheap -- these groups are small (tens of rows), not thousands.

    autojunk=False is deliberate: SequenceMatcher's default autojunk=True
    treats characters that occur too often as "popular" and downweights
    them, but only for the second sequence -- for two paragraph-length
    strings over 200 chars this makes ratio() direction-dependent (a real
    case: ratio(a, b) came out 0.94, ratio(b, a) came out 0.76 for the same
    pair), so a duplicate could be missed or not depending on which side of
    the comparison it landed on. Off, ratio() is symmetric and appropriate
    for comparing prose rather than source-code-like repetitive text."""

    # The key includes scope_type, city_name and tier_scope, not just
    # country/state/section. A bullet that names several cities is expanded
    # into one row per city with the SAME description (that granularity is
    # deliberate in the flat template). Comparing descriptions alone treated
    # those city rows as duplicates of each other and collapsed them: on a
    # full India run it removed 106 rows, 104 of which were city/tier
    # siblings and only 2 true near duplicates (Karnataka lost 51 rows,
    # Uttar Pradesh 16, Tamil Nadu 17). Rows are now only compared against
    # others for the same location.
    groups = {}

    for r in rows:
        key = (
            r.get("country_name"), r.get("state_name"), r.get("section_name"),
            r.get("scope_type"), r.get("city_name"), r.get("tier_scope"),
        )
        groups.setdefault(key, []).append(r)

    out = []

    for group_rows in groups.values():

        # Same result as comparing every pair with
        # SequenceMatcher(None, desc, t, autojunk=False).ratio(), just
        # faster: each kept text gets one matcher (kept text as seq2, the
        # expensive side to index, each new description set as seq1: the
        # same (a, b) orientation), and the exact ratio() is only computed
        # for pairs that pass two upper bounds. real_quick_ratio() is a
        # documented upper bound on ratio(); so is 2 * LCS / (len a + len b),
        # because SequenceMatcher's matching blocks always form a common
        # subsequence, so what it matches can never exceed the LCS. A pair
        # failing either bound can never reach the threshold. Without the
        # bounds, 3000 character items (the split size for oversized PDF
        # blocks) cost about 0.02 to 0.06 s per pair: 78 s for the 148 rows
        # of one country.
        kept = []

        for r in group_rows:

            desc = (r.get("description") or "").strip().lower()

            is_dupe = False

            for sm, masks, n in kept:

                total = len(desc) + n

                if total and 2.0 * min(len(desc), n) / total < threshold:
                    continue

                if total and 2.0 * _lcs_length(masks, n, desc) / total < threshold:
                    continue

                sm.set_seq1(desc)

                if sm.ratio() >= threshold:
                    is_dupe = True
                    break

            if is_dupe:
                continue

            sm = difflib.SequenceMatcher(None, autojunk=False)
            sm.set_seq2(desc)
            kept.append((sm, _lcs_masks(desc), len(desc)))
            out.append(r)

    return out


def extract_india_rows_from_workbook(wb, source_dates_index=None):

    if "India-State wise" not in wb.sheetnames:
        return {}

    ws = wb["India-State wise"]

    # Header states in columns 3+
    states = []

    for c in range(3, ws.max_column + 1):

        v = ws.cell(1, c).value

        if v:
            states.append((c, str(v).strip()))

    # Validity row
    validity_by_col = {
        c: str(ws.cell(2, c).value or "").strip()
        for c, _ in states
    }

    incentive_by_row = _extract_incentive_type_map(ws)

    india_rows_by_state = {state: [] for _, state in states}

    for r in range(3, ws.max_row + 1):

        incentive_type = incentive_by_row.get(r, "")

        scope_label = str(ws.cell(r, 2).value or "").strip()

        if not incentive_type or not scope_label:
            continue

        section = SECTION_MAP.get(incentive_type)

        if not section:
            continue

        scope_type, tier_scope = SCOPE_MAP_INDIA.get(
            scope_label, ("State", scope_label)
        )

        for c, state in states:

            bullets = parse_bullets(ws.cell(r, c).value)

            if not bullets:
                continue

            validity_start, vs_prec, validity_end, ve_prec = parse_validity(
                validity_by_col.get(c, "")
            )

            for b in bullets:

                text_dates = extract_validity_from_text(b)

                start, s_prec, end, e_prec = _apply_source_fallback(
                    text_dates,
                    (source_dates_index or {}).get(("India", state)),
                    (validity_start, vs_prec, validity_end, ve_prec),
                )

                # Only exclude on a date that's actually about this bullet:
                # its own text, or the manually-curated Validity column for
                # this state. A full-document source scan can find ONE
                # stray date anywhere on a multi-topic page and, if used for
                # exclusion, would wrongly blank out every other bullet from
                # that source. It's fine as a display fallback (folded into
                # `end` above), just not as grounds to drop content.
                if is_expired(text_dates[2] or validity_end, text_dates[3] or ve_prec):
                    continue

                cities = detect_cities(
                    b, INDIAN_STATE_CITIES.get(state, {})
                )

                relevance = compute_gcc_relevance(b, "India")

                if cities:
                    for city_name, tier in cities:
                        india_rows_by_state[state].append({
                            "section_name": section,
                            "scope_type": "City",
                            "country_name": "India",
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
                    india_rows_by_state[state].append({
                        "section_name": section,
                        "scope_type": scope_type,
                        "country_name": "India",
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

    for state in india_rows_by_state:
        india_rows_by_state[state] = _fuzzy_dedupe_rows(india_rows_by_state[state])

    return india_rows_by_state


def extract_global_rows_from_workbook(wb, source_dates_index=None):

    skip_sheets = {"Definition & Methodology", "Summary", "India-State wise",
                   "Country Template", "Countries"}

    global_rows_by_country = {}

    for sn in wb.sheetnames:

        if sn in skip_sheets:
            continue

        ws = wb[sn]

        if ws.max_column < 3 or ws.max_row < 3:
            continue

        country_header = str(ws.cell(1, 3).value or sn).strip()

        validity_str = str(ws.cell(2, 3).value or "").strip()

        incentive_by_row = _extract_incentive_type_map(ws)

        rows = []

        for r in range(3, ws.max_row + 1):

            incentive_type = incentive_by_row.get(r, "")

            scope_label = str(ws.cell(r, 2).value or "").strip()

            if not incentive_type or not scope_label:
                continue

            section = SECTION_MAP.get(incentive_type)

            if not section:
                continue

            scope_type, tier_scope = SCOPE_MAP_GLOBAL.get(
                scope_label, ("Country", scope_label)
            )

            bullets = parse_bullets(ws.cell(r, 3).value)

            if not bullets:
                continue

            validity_start, vs_prec, validity_end, ve_prec = parse_validity(
                validity_str
            )

            for b in bullets:

                text_dates = extract_validity_from_text(b)

                start, s_prec, end, e_prec = _apply_source_fallback(
                    text_dates,
                    (source_dates_index or {}).get(("Global", country_header)),
                    (validity_start, vs_prec, validity_end, ve_prec),
                )

                # See the matching comment in extract_india_rows_from_workbook:
                # only a date tied to this bullet (its own text) or the
                # manually-entered Validity column should gate exclusion. A
                # source-level full-document scan is a display fallback only.
                if is_expired(text_dates[2] or validity_end, text_dates[3] or ve_prec):
                    continue

                cities = detect_cities(
                    b, COUNTRY_CITIES.get(country_header, {})
                )

                relevance = compute_gcc_relevance(b, country_header)

                if cities:
                    for city_name, tier in cities:
                        rows.append({
                            "section_name": section,
                            "scope_type": "City",
                            "country_name": country_header,
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
                        "country_name": country_header,
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

        if rows:
            global_rows_by_country[country_header] = _fuzzy_dedupe_rows(rows)

    return global_rows_by_country


def extract_tamil_nadu_from_pdf():
    """Extract Tamil Nadu content directly from the local PDF the user placed."""

    if not os.path.exists(TAMIL_NADU_PDF):

        print(f"⚠ {TAMIL_NADU_PDF} not found; Tamil Nadu will not be added from PDF.")

        return []

    print(f"→ Extracting Tamil Nadu from {TAMIL_NADU_PDF}")

    paragraphs = PolicyPDFReader.read(TAMIL_NADU_PDF)

    buckets = PolicyBucketer.bucket(paragraphs)

    validity_start, vs_prec, validity_end, ve_prec = parse_validity(
        TAMIL_NADU_VALIDITY
    )

    rows = []

    for incentive_type, scopes in buckets.items():

        section = SECTION_MAP.get(incentive_type)

        if not section:
            continue

        for scope_label, items in scopes.items():

            if not items:
                continue

            scope_type, tier_scope = SCOPE_MAP_INDIA.get(
                scope_label, ("State", scope_label)
            )

            for b in items:

                text_dates = extract_validity_from_text(b)

                start, s_prec, end, e_prec = _apply_source_fallback(
                    text_dates, None,
                    (validity_start, vs_prec, validity_end, ve_prec),
                )

                if is_expired(end, e_prec):
                    continue

                cities = detect_cities(
                    b, INDIAN_STATE_CITIES.get("Tamil Nadu", {})
                )

                relevance = compute_gcc_relevance(b, "India")

                if cities:
                    for city_name, tier in cities:
                        rows.append({
                            "section_name": section,
                            "scope_type": "City",
                            "country_name": "India",
                            "state_name": "Tamil Nadu",
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
                        "country_name": "India",
                        "state_name": "Tamil Nadu",
                        "city_name": "",
                        "tier_scope": tier_scope,
                        "description": b,
                        "validity_start_date": start,
                        "validity_start_precision": s_prec,
                        "validity_end_date": end,
                        "validity_end_precision": e_prec,
                        "gcc_relevance": relevance,
                    })

    print(f"  Tamil Nadu rows generated: {len(rows)}")

    return rows


def write_new_template(india_rows, global_rows, output_path):

    wb = openpyxl.Workbook()

    wb.remove(wb.active)

    _write_flat_sheet(wb, "India", india_rows)

    # Split Global into signal (score >= 1) and noise (score == 0) sheets.
    # This keeps every row in the workbook while making the main Global
    # sheet reviewable at a glance.
    signal_rows = [
        r for r in global_rows
        if int(r.get("gcc_relevance") or 0) >= 1
    ]

    other_rows = [
        r for r in global_rows
        if int(r.get("gcc_relevance") or 0) == 0
    ]

    _write_flat_sheet(wb, "Global", signal_rows)
    _write_flat_sheet(wb, "Other - Global", other_rows)

    wb.save(output_path)


def _clean_cell_value(value):
    """Strip Excel-illegal control characters and cap length. Extracted PDF
    text occasionally carries stray control bytes (confirmed on a Maryland
    incentives PDF) that raise IllegalCharacterError deep inside openpyxl if
    written as-is -- this mirrors the same guard exporter/policy_excel_
    exporter.py's wide-sheet writer already has, applied here so the flat
    sheet (India/Global/every state-level target) can't hit it either."""

    if not isinstance(value, str):
        return value

    cleaned = ILLEGAL_CHARACTERS_RE.sub("", value)

    if len(cleaned) > 32000:
        cleaned = cleaned[:32000] + "…[truncated]"

    return cleaned


def _write_flat_sheet(wb, sheet_name, rows):

    ws = wb.create_sheet(sheet_name)

    for c, h in enumerate(HEADERS, start=1):

        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(
            wrap_text=True, vertical="center", horizontal="center"
        )
        cell.border = BORDER

    for r_idx, row in enumerate(rows, start=2):

        for c_idx, header in enumerate(HEADERS, start=1):

            cell = ws.cell(row=r_idx, column=c_idx, value=_clean_cell_value(row.get(header, "")))
            cell.alignment = WRAP
            cell.border = BORDER

    widths = {
        "A": 32, "B": 12, "C": 16, "D": 16, "E": 14, "F": 22,
        "G": 90, "H": 14, "I": 12, "J": 14, "K": 12, "L": 12,
    }

    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.freeze_panes = "A2"

    ws.row_dimensions[1].height = 30

    for r in range(2, len(rows) + 2):
        ws.row_dimensions[r].height = 55


def run():
    """Callable entry point. Returns the output file path, or None on failure."""

    latest = find_latest_workbook()

    if latest is None:

        print("✗ No output/gcc_incentives_global_*.xlsx found. Run main_global.py first.")

        return None

    print(f"→ Reading latest workbook: {os.path.basename(latest)}")

    wb = openpyxl.load_workbook(latest, data_only=True)

    # Build a per-jurisdiction fallback of validity dates by scanning the
    # actual source files (not just the workbook cells).
    source_dates_index = {}
    if os.path.exists(SOURCES_FILE):
        try:
            srcs = SourcesLoader.load(SOURCES_FILE)
            source_dates_index = build_source_dates_index(srcs)
        except Exception as exc:
            print(f"  ⚠ source date scan skipped: {exc}")

    india_by_state = extract_india_rows_from_workbook(
        wb, source_dates_index=source_dates_index,
    )

    global_by_country = extract_global_rows_from_workbook(
        wb, source_dates_index=source_dates_index,
    )

    # Tamil Nadu: use the workbook (which is paraphrased when the pipeline
    # ran with --paraphrase). Only fall back to raw PDF extraction when
    # the workbook has nothing for Tamil Nadu.
    if not india_by_state.get("Tamil Nadu"):

        tn_rows = extract_tamil_nadu_from_pdf()

        if tn_rows:

            india_by_state["Tamil Nadu"] = tn_rows

            print("  ⓘ Tamil Nadu filled from raw PDF (workbook had none, "
                  "content is NOT paraphrased)")

    else:

        print(f"  ⓘ Tamil Nadu content taken from workbook "
              f"({len(india_by_state['Tamil Nadu'])} rows, paraphrased if "
              "the pipeline was paraphrased)")

    # Sort each jurisdiction's rows by GCC relevance (desc) so the strongest
    # signal appears first within each state/country block.
    india_rows = []
    for state, rows in india_by_state.items():
        rows.sort(key=lambda r: -int(r.get("gcc_relevance") or 0))
        india_rows.extend(rows)

    global_rows = []
    for country, rows in global_by_country.items():
        rows.sort(key=lambda r: -int(r.get("gcc_relevance") or 0))
        global_rows.extend(rows)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    today = date.today().isoformat()

    output_path = os.path.join(
        OUTPUT_DIR, f"new_template_GCC_policy_{today}.xlsx"
    )

    write_new_template(india_rows, global_rows, output_path)

    signal_count = sum(
        1 for r in global_rows if int(r.get("gcc_relevance") or 0) >= 1
    )
    other_count = len(global_rows) - signal_count

    print(f"\n✓ Wrote {output_path}")
    print(f"  India rows: {len(india_rows)} across {len(india_by_state)} states")
    print(f"  Global rows: {signal_count} (score >= 1) in 'Global' sheet")
    print(f"  Other Global: {other_count} (score = 0) in 'Other - Global' sheet")

    return output_path


def main():
    run()


if __name__ == "__main__":
    main()
