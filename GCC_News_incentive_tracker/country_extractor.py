"""Single entry point for the whole GCC incentive-tracking pipeline.

Replaces the previously separate main_usa.py, main_country_states.py, and
main_global.py, which had drifted into three copies of nearly the same
extract -> translate -> bucket -> paraphrase -> flatten pipeline (and, in
main_usa.py and main_global.py's case, an older bucket-before-translate
order that silently dropped most non-English content -- see the
build_state_data() docstring below).

Targets:
  United States, United Kingdom, Germany, France, Japan
      -> state/region/prefecture-level workbooks (COUNTRY_STATES below)
  Global
      -> the ~330-country Global sheet plus India's GCC-policy states

Default (no flags) runs every target, one after another. Pass one or more
per-country flags (--usa, --uk, --germany, --france, --japan, --global,
--india) to run only those. --global builds India's 8 states together with
the ~330-country Global sheet; --india refreshes just those 8 states (own
old_template/India and new_template/India folders) without a full Global
run. The default run doesn't list India separately since Global covers it.

Output layout (see project root for the sibling folders this reads/writes):
  output/old_template/<Target>/...   wide, audit-trail workbooks
  output/new_template/<Target>/...   flat, GCC-relevance-scored workbooks
  input/sources/sources*.xlsx        per-target source-list workbooks

Run:
  python3 country_extractor.py                  # every target
  python3 country_extractor.py --germany         # one target
  python3 country_extractor.py --germany --france  # more than one
  python3 country_extractor.py --global --refresh

Long runs (a task time limit, a laptop that sleeps):
  python3 country_extractor.py --global --countries "Brazil,Spain" \\
        --resume --max-minutes 8
Stops cleanly after about 8 minutes at the next paraphrase chunk, with the
paraphrase cache saved (it is also saved every 8 items regardless) and, with
--resume, the fetched and translated text kept. Exit code 3 means "stopped at
the time limit, run the same command again"; it continues instead of starting
over, and the saved text is deleted once the run completes.
"""

import argparse
import glob
import hashlib
import os
import pickle
import re
import sys
import threading
import time
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font

from feeds.policy_downloader import PolicyDownloader
from feeds.previous_workbook_reader import PreviousWorkbookReader
from feeds.sources_loader import SourcesLoader
from feeds.template_loader import TemplateLoader
from extractor.policy_pdf_reader import PolicyPDFReader
from extractor.policy_html_reader import PolicyHTMLReader
from extractor.policy_bucketer import PolicyBucketer
from extractor.paraphraser import OllamaParaphraser, ParaphraseDeadline
from extractor.translator import ArgosTranslator
from exporter.policy_excel_exporter import PolicyExcelExporter
from exporter.sources_excel_exporter import SourcesExcelExporter

# Reused as a library of pure helper functions (date extraction, city
# detection, relevance scoring, flat-sheet writer, and the India/Global
# flatten step). Importing this does not execute its own pipeline.
import new_template as nt


# ---------------------------------------------------------------------------
# Paths -- the old_template / new_template / sources layout.
# ---------------------------------------------------------------------------

SOURCES_DIR = os.path.join("input", "sources")     # sources*.xlsx workbooks
TEMPLATE_FILE = os.path.join("input", "template.xlsx")  # Global's ~330-country list

OLD_TEMPLATE_ROOT = os.path.join("output", "old_template")
NEW_TEMPLATE_ROOT = os.path.join("output", "new_template")

INDIA_SOURCES_DIR = os.path.join("sources", "india")   # raw PDF/HTML cache
GLOBAL_SOURCES_DIR = os.path.join("sources", "global")  # raw PDF/HTML cache

DOWNLOAD_WORKERS = 8
EXTRACT_WORKERS = 6
SCAN_WORKERS = int(os.environ.get("SCAN_WORKERS", "10"))


def slug(name):
    return name.lower().replace(" ", "_").replace("-", "_").replace("'", "")


def empty_buckets():

    return {
        it: {scope: [] for scope in PolicyBucketer.SCOPES}
        for it in PolicyBucketer.INCENTIVE_TYPES
    }


def _dedupe_items(items):
    """Drop near-duplicate bullets within one incentive-type/scope bucket --
    e.g. the same paragraph reached via both a PDF and a linked HTML page,
    or picked up twice by the depth=1 crawl. Keys on the first 60 characters,
    the same heuristic PolicyExcelExporter._dedupe() already uses for the
    wide-sheet writer; applied here too so the flat/new_template output
    (which doesn't go through that writer) can't carry duplicate rows."""

    seen = set()
    out = []

    for it in items:

        key = it[:60]

        if key in seen:
            continue

        seen.add(key)
        out.append(it)

    return out


@lru_cache(maxsize=None)
def _country_name_re(name_lower):
    return re.compile(rf"\b{re.escape(name_lower)}\b")


def _names_country(name_lower, text_lower):
    """Whole word match. Plain substring matching misfired on common words:
    "cuba" inside "incubator"/"incubation" (the single most common false hit,
    since incubators and accelerators are standard innovation incentive
    language), "oman" inside "woman" and "Romania", "india" inside "Indiana".
    A genuine row that said "incubator" without naming its own country was
    treated as being about Cuba and dropped by both cross country filters."""
    return bool(_country_name_re(name_lower).search(text_lower))


def filter_cross_country_bleed_rows(rows, known_countries):
    """Row-level second pass of filter_cross_country_bleed(), applied to the
    FINAL description text after paraphrasing.

    filter_cross_country_bleed() runs pre-paraphrase and catches raw content
    that was already about a different country (Spain's "Comparison of
    national incentives in Europe" case). It can't catch paraphrasing itself
    introducing or swapping a country name -- confirmed case: a Thailand
    bullet listing real Thai province names (Kalasin, Chaiyaphum, Nakhon
    Phanom, ...) came back from Ollama saying "incentives for businesses
    investing in Vietnam" instead of Thailand, and a South Korea bullet came
    back entirely about a Philippines tax program. Same conservative trigger:
    drop a row only when its own country isn't named in the description AND
    a different known country clearly is."""

    lowered = {c: c.lower() for c in known_countries}

    kept = []

    for r in rows:

        country = r.get("country_name") or ""
        own = country.lower()
        desc = (r.get("description") or "").lower()

        if own and not _names_country(own, desc):

            other_hit = any(
                _names_country(oc_lower, desc)
                for oc, oc_lower in lowered.items()
                if oc != country
            )

            if other_hit:
                continue

        kept.append(r)

    return kept


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


# ---------------------------------------------------------------------------
# Per-country configuration for the state/region/prefecture-level targets:
# the official list of jurisdictions and a hand-curated Tier 1 / Tier 2 & 3
# city registry for each, at the same level of detail as new_template.py's
# COUNTRY_CITIES entries. United States uses the same 50-state list and city
# registry main_usa.py did.
# ---------------------------------------------------------------------------

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

COUNTRY_STATES = {
    "United States": US_STATES,
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
    "United States": US_STATE_CITIES,
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
    return os.path.join(SOURCES_DIR, f"sources_{slug(country)}.xlsx")


def raw_cache_dir_for(country):
    return os.path.join("sources", slug(country))


def old_template_dir_for(target):
    return os.path.join(OLD_TEMPLATE_ROOT, target)


def new_template_dir_for(target):
    return os.path.join(NEW_TEMPLATE_ROOT, target)


# ---------------------------------------------------------------------------
# Shared building blocks for the state/region/prefecture-level targets.
# ---------------------------------------------------------------------------

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
         os.path.join(raw_cache_dir_for(country), slug(state), filename), force)
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


def _read_source(jurisdiction, kind, ref, reader):
    """Read one configured source and say so when it yields nothing.

    A failed fetch used to look exactly like an empty source: a blocked,
    dead or truncated source silently lowered a country's row count, and
    the drop was only visible weeks later by comparing output files."""

    try:
        paragraphs = reader()
    except Exception as exc:
        paragraphs = []
        reason = f"error {type(exc).__name__}"
    else:
        reason = None

    if not paragraphs:

        if reason is None:

            if kind == "PDF":
                reason = ("file missing, truncated or image only"
                          if not PolicyDownloader._is_complete_pdf(ref)
                          else "no text extracted")
            else:
                reason = PolicyHTMLReader.FAILURES.get(ref, "page fetched but no text extracted")

        print(f"  ⚠ {jurisdiction}: {kind} returned no paragraphs ({reason}): "
              f"{os.path.basename(ref) if kind == 'PDF' else ref}")

    return paragraphs


def _extract_one_state(country, state, paths, html_urls):

    state_dir = os.path.join(raw_cache_dir_for(country), slug(state))

    paragraphs = []

    for path in paths:
        paragraphs.extend(_read_source(state, "PDF", path, lambda p=path: PolicyPDFReader.read(p)))

    for url in html_urls:
        paragraphs.extend(_read_source(
            state, "HTML", url,
            lambda u=url: PolicyHTMLReader.read(u, country_dir=state_dir, depth=1, pdf_reader=PolicyPDFReader),
        ))

    return (state, paragraphs)


# ---------------------------------------------------------------------------
# --resume: keep the expensive, network dependent part of a run (fetch,
# translate, bucket) so an interrupted run does not start from zero.
# ---------------------------------------------------------------------------

RESUME = False
RESUME_DIR = os.path.join("output", ".resume")
RESUME_MAX_AGE_HOURS = 24
_resume_files = set()


def _resume_path(kind, key):
    return os.path.join(RESUME_DIR, f"{kind}_{hashlib.sha1(repr(key).encode('utf-8')).hexdigest()[:16]}.pkl")


def _resume_load(path):
    """Saved extraction from an earlier interrupted run, or None. Ignored
    after RESUME_MAX_AGE_HOURS so stale page content is never reused."""

    if not RESUME or not os.path.exists(path):
        return None

    if time.time() - os.path.getmtime(path) > RESUME_MAX_AGE_HOURS * 3600:
        return None

    try:
        with open(path, "rb") as f:
            value = pickle.load(f)
    except Exception:
        return None

    _resume_files.add(path)

    return value


def _resume_save(path, value):

    if not RESUME:
        return

    os.makedirs(RESUME_DIR, exist_ok=True)

    tmp_path = path + ".part"

    with open(tmp_path, "wb") as f:
        pickle.dump(value, f)

    os.replace(tmp_path, path)

    _resume_files.add(path)


def _resume_clear():
    """Delete the saved extraction once a target has finished, so the next
    ordinary run fetches fresh content."""

    for path in list(_resume_files):
        try:
            os.remove(path)
        except OSError:
            pass

    _resume_files.clear()


def build_state_data(country, state_available, state_html, state_validity, states,
                      translator=None):

    resume_path = _resume_path("states", (country, tuple(states)))

    saved = _resume_load(resume_path)

    if saved is not None:
        print(f"↻ resume: reusing the extraction saved by an earlier run ({country}, {len(states)} states)")
        return saved

    result = _build_state_data(country, state_available, state_html, state_validity, states, translator)

    _resume_save(resume_path, result)

    return result


def _build_state_data(country, state_available, state_html, state_validity, states,
                       translator=None):
    """Extract raw paragraphs per state, translate to English, THEN bucket.

    Bucketing must happen after translation, not before: PolicyBucketer's
    keyword vocabulary is English/Spanish only, so feeding it raw German/
    French/Japanese text drops almost everything at the classification
    step. Confirmed directly on a real source (Hesse's WIBank page: 314
    extracted paragraphs, 0 bucketed items under the old bucket-then-
    translate order; 32 bucketed items once translated first)."""

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

        for incentive_type, scopes in data.items():
            for scope, items in scopes.items():
                scopes[scope] = _dedupe_items(items)

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
                raw_cache_dir_for(country), slug(state), filename), False))

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

                    # Only exclude on a date tied to this bullet: its own
                    # text, or the manually-entered Validity column for this
                    # state. A full-document source scan can find ONE stray
                    # date anywhere on a multi-topic page and, used for
                    # exclusion, would wrongly blank out every other bullet
                    # from that source -- confirmed on France/Corse,
                    # Japan/Osaka, and Germany/North Rhine-Westphalia, each
                    # wiped out entirely by one unrelated page date. It's
                    # fine as a display fallback (folded into `end` above),
                    # just not as grounds to drop content.
                    if nt.is_expired(text_dates[2] or validity_end, text_dates[3] or ve_prec):
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

    for state in rows_by_state:
        rows_by_state[state] = nt._fuzzy_dedupe_rows(rows_by_state[state])

    return rows_by_state


def build_state_sources_rows(country, srcs, state_available):
    """Audit-trail rows: which PDF/HTML sources were configured for this
    country, whether each was actually downloaded/used, and where the local
    copy lives. Mirrors main_usa.py's build_sources_rows() -- main_country_
    states.py (the UK/Germany/France/Japan predecessor of this pipeline)
    never had this step, so United Kingdom, Germany, France, and Japan have
    never had a gcc_sources_*.xlsx audit file, and USA's is stale (last
    written by the old main_usa.py, not by this merged pipeline)."""

    rows = []

    for state, pdfs in srcs["country_pdfs"].items():

        downloaded = {os.path.basename(p) for p in state_available.get(state, [])}

        for filename, url, _ in pdfs:

            is_local = not url
            status = "Used" if filename in downloaded else "Failed"
            local = (
                os.path.join(raw_cache_dir_for(country), slug(state), filename)
                if filename in downloaded else ""
            )

            if is_local:
                note = "Local file, extracted" if status == "Used" else "Local file missing at expected path"
            else:
                note = "Downloaded and extracted" if status == "Used" else "Configured but download failed"

            rows.append((country, state, url or "(local file)", "PDF", status, local, note))

    for state, urls in srcs["country_html"].items():
        for url in urls:
            rows.append((country, state, url, "HTML", "Scraped", "", "HTML scraped and bucketed"))

    return rows


# ---------------------------------------------------------------------------
# Excel export for the state-level targets: reuses PolicyExcelExporter's
# diff/style building blocks.
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


def write_state_level_outputs(country, state_data, validity_by_state, rows_by_state,
                               write_old_template=True):

    new_dir = new_template_dir_for(country)
    os.makedirs(new_dir, exist_ok=True)

    today = date.today().isoformat()

    wide_path = None

    if write_old_template:

        old_dir = old_template_dir_for(country)
        os.makedirs(old_dir, exist_ok=True)

        wide_path = os.path.join(old_dir, f"gcc_incentives_{slug(country)}_states_{today}.xlsx")

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

    flat_path = os.path.join(new_dir, f"new_template_{slug(country)}_states_policy_{today}.xlsx")

    wb2 = openpyxl.Workbook()
    wb2.remove(wb2.active)
    nt._write_flat_sheet(wb2, country[:31], signal_rows)
    nt._write_flat_sheet(wb2, f"Other - {country}"[:31], other_rows)
    wb2.save(flat_path)

    return wide_path, flat_path, len(signal_rows), len(other_rows)


def run_state_level_target(country, args):

    print(f"\n{'=' * 70}\nTarget: {country}\n{'=' * 70}\n")

    sources_file = sources_file_for(country)

    if not os.path.exists(sources_file):
        print(f"✗ {sources_file} not found. Skipping {country}.")
        return

    print(f"Loading {country} sources from {sources_file}\n")

    srcs = SourcesLoader.load(sources_file)

    states = list(srcs["country_pdfs"].keys())
    for s in srcs["country_html"]:
        if s not in states:
            states.append(s)
    for s in COUNTRY_STATES.get(country, []):
        if s not in states:
            states.append(s)

    os.makedirs(raw_cache_dir_for(country), exist_ok=True)

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
            print("⚠ Ollama not reachable/model not pulled; skipping paraphrasing.")

    source_dates_index = build_source_dates_index(country, srcs)

    rows_by_state = extract_state_rows(country, state_data, validity_by_state, source_dates_index)

    write_old_template = not getattr(args, "no_old_template", False)

    wide_path, flat_path, signal_n, other_n = write_state_level_outputs(
        country, state_data, validity_by_state, rows_by_state,
        write_old_template=write_old_template,
    )

    print(f"\n✓ Wrote {flat_path}  ({country}: {signal_n} rows, Other - {country}: {other_n} rows)")

    if write_old_template:

        print(f"✓ Wrote {wide_path}")

        sources_rows = build_state_sources_rows(country, srcs, state_available)

        sources_path = os.path.join(
            old_template_dir_for(country), f"gcc_sources_{slug(country)}_{date.today().isoformat()}.xlsx",
        )

        SourcesExcelExporter.export(sources_rows, sources_path)

        print(f"✓ Wrote {sources_path}  ({len(sources_rows)} source rows)")
    else:
        print("  (--no-old-template: wide workbook and gcc_sources file skipped)")


# ---------------------------------------------------------------------------
# Global + India target: the ~330-country Global sheet plus India's 8
# GCC-policy states, combined into one wide workbook (like the state-level
# targets, this now translates each jurisdiction's raw paragraphs BEFORE
# bucketing, and BOTH the India and Global row-extraction in new_template.py
# only let a bullet's own text -- not a source-wide date scan -- exclude it
# as expired).
# ---------------------------------------------------------------------------

def _extract_one_jurisdiction(jurisdiction, paths, html_urls, jurisdiction_dir):

    paragraphs = []

    for path in paths:
        paragraphs.extend(_read_source(jurisdiction, "PDF", path, lambda p=path: PolicyPDFReader.read(p)))

    for url in html_urls:
        paragraphs.extend(_read_source(
            jurisdiction, "HTML", url,
            lambda u=url: PolicyHTMLReader.read(u, country_dir=jurisdiction_dir, depth=1, pdf_reader=PolicyPDFReader),
        ))

    return (jurisdiction, paragraphs)


def filter_cross_country_bleed(country_data, known_countries):
    """Drop bucketed items that are actually about a DIFFERENT country than
    the one they're filed under.

    Confirmed case: Spain's own official incentives page (a legitimate
    Spain source) links to "Comparison of national incentives in Europe",
    a PDF that discusses Germany, France, Poland, Hungary and the
    Netherlands. The depth=1 crawl follows that link as a normal resource,
    so every bullet from it -- including ones entirely about Poland or
    Hungary -- gets extracted under Spain's bucket, since extraction has no
    concept of "this paragraph is actually about a different country."

    Conservative trigger: drop an item only when its own country isn't
    named in the text AND a different known country clearly is. Country-
    specific text that simply doesn't repeat the country's name (the much
    more common case) mentions no OTHER country either, so it's untouched."""

    lowered = {c: c.lower() for c in known_countries}

    for country, buckets in country_data.items():

        own = country.lower()

        for incentive_type, scopes in buckets.items():

            for scope, items in scopes.items():

                kept = []

                for text in items:

                    low = text.lower()

                    if not _names_country(own, low):

                        other_hit = any(
                            _names_country(oc_lower, low)
                            for oc, oc_lower in lowered.items()
                            if oc != country
                        )

                        if other_hit:
                            continue

                    kept.append(text)

                scopes[scope] = kept

    return country_data


def ensure_pdf_sources_global(india_pdfs, country_pdfs, force=False):

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
         os.path.join(GLOBAL_SOURCES_DIR, slug(country), filename), force)
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


def _build_jurisdiction_data(available, html, jurisdictions, cache_root, translator, label):

    resume_path = _resume_path("jurisdictions", (label, tuple(jurisdictions)))

    saved = _resume_load(resume_path)

    if saved is not None:
        print(f"↻ resume: reusing the extraction saved by an earlier run ({label}, {len(jurisdictions)} jurisdictions)")
        return saved

    result = _build_jurisdiction_data_uncached(available, html, jurisdictions, cache_root, translator, label)

    _resume_save(resume_path, result)

    return result


def _build_jurisdiction_data_uncached(available, html, jurisdictions, cache_root, translator, label):
    """Shared by India (8 states) and Global (~330 countries): extract raw
    paragraphs per jurisdiction in parallel, translate, THEN bucket -- same
    fix as build_state_data() above, same reason (PolicyBucketer's keyword
    vocabulary is English/Spanish only)."""

    tasks = [
        (j, available.get(j, []), html.get(j, []), os.path.join(cache_root, slug(j)))
        for j in jurisdictions
    ]

    raw_by_jurisdiction = {}

    with ThreadPoolExecutor(max_workers=EXTRACT_WORKERS) as ex:

        futures = {ex.submit(_extract_one_jurisdiction, *t): t[0] for t in tasks}

        for f in as_completed(futures):
            j, paragraphs = f.result()
            raw_by_jurisdiction[j] = paragraphs

    data = {}

    for j, paragraphs in raw_by_jurisdiction.items():

        if not paragraphs:
            data[j] = empty_buckets()
            print(f"→ {j}: bucketed 0 items from 0 paragraphs")
            continue

        texts = [p["text"] for p in paragraphs]

        if translator is not None:
            texts = translator.translate_many(texts, label=f"{label}/{j[:14]}/raw")

        translated_paragraphs = [
            {"heading": p.get("heading", ""), "text": t}
            for p, t in zip(paragraphs, texts)
        ]

        bucketed = PolicyBucketer.bucket(translated_paragraphs)

        for incentive_type, scopes in bucketed.items():
            for scope, items in scopes.items():
                scopes[scope] = _dedupe_items(items)

        data[j] = bucketed

        total = sum(len(items) for scopes in bucketed.values() for items in scopes.values())

        print(f"→ {j}: bucketed {total} items from {len(paragraphs)} paragraphs")

    return data


def build_global_sources_rows(srcs, india_available, country_available):

    rows = []

    for state, pdfs in srcs["india_pdfs"].items():

        downloaded = {os.path.basename(p) for p in india_available.get(state, [])}

        for filename, url, _ in pdfs:

            is_local = not url
            status = "Used" if filename in downloaded else "Failed"
            local = os.path.join(INDIA_SOURCES_DIR, filename) if filename in downloaded else ""

            if is_local:
                note = "Local file, extracted" if status == "Used" else "Local file missing at expected path"
            else:
                note = "Downloaded and extracted" if status == "Used" else "Configured but download failed"

            rows.append(("India", state, url or "(local file)", "PDF", status, local, note))

    for state, urls in srcs["india_html"].items():
        for url in urls:
            rows.append(("India", state, url, "HTML", "Scraped", "", "HTML scraped and bucketed"))

    for country, pdfs in srcs["country_pdfs"].items():

        downloaded = {os.path.basename(p) for p in country_available.get(country, [])}

        for filename, url, _ in pdfs:

            is_local = not url
            status = "Used" if filename in downloaded else "Failed"
            local = (
                os.path.join(GLOBAL_SOURCES_DIR, slug(country), filename)
                if filename in downloaded else ""
            )

            if is_local:
                note = "Local file, extracted" if status == "Used" else "Local file missing at expected path"
            else:
                note = "Downloaded and extracted" if status == "Used" else "Configured but download failed"

            rows.append(("Global", country, url or "(local file)", "PDF", status, local, note))

    for country, urls in srcs["country_html"].items():
        for url in urls:
            rows.append(("Global", country, url, "HTML", "Scraped", "", "HTML scraped and bucketed"))

    return rows


def extract_global_rows_from_data(country_data, validity_by_country, source_dates_index,
                                   known_countries=None):
    """In-memory equivalent of new_template.extract_global_rows_from_workbook(),
    built for --no-old-template runs: that function reads bullets back out of
    the saved wide workbook (parsing "* " joined cell text via parse_bullets),
    which only exists if the wide file was written. This reads the same
    country_data dict build_global_target already has in memory, before any
    wide-format write, so a subset/no-old-template run doesn't need one."""

    rows_by_country = {c: [] for c in country_data}

    for country, buckets in country_data.items():

        validity_start, vs_prec, validity_end, ve_prec = nt.parse_validity(
            validity_by_country.get(country, "")
        )

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
                        text_dates,
                        source_dates_index.get(("Global", country)),
                        (validity_start, vs_prec, validity_end, ve_prec),
                    )

                    # Same text-only expiry gate as extract_state_rows() and
                    # new_template's own India/Global extraction: a source-
                    # wide date scan is a display fallback only, never
                    # grounds by itself to drop a bullet.
                    if nt.is_expired(text_dates[2] or validity_end, text_dates[3] or ve_prec):
                        continue

                    cities = nt.detect_cities(b, nt.COUNTRY_CITIES.get(country, {}))

                    relevance = nt.compute_gcc_relevance(b, country)

                    if cities:
                        for city_name, tier in cities:
                            rows_by_country[country].append({
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
                        rows_by_country[country].append({
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

    for country in rows_by_country:

        rows = nt._fuzzy_dedupe_rows(rows_by_country[country])

        if known_countries:
            rows = filter_cross_country_bleed_rows(rows, known_countries)

        rows_by_country[country] = rows

    return rows_by_country


def extract_india_rows_from_data(state_data, validity_by_state, source_dates_index):
    """In-memory equivalent of new_template.extract_india_rows_from_workbook(),
    for the India-only target: same validity/expiry/city/relevance logic, but
    read straight from the bucketed state_data instead of re-parsing a saved
    wide workbook, so an India-only run doesn't need one."""

    rows_by_state = {s: [] for s in state_data}

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
                        source_dates_index.get(("India", state)),
                        (validity_start, vs_prec, validity_end, ve_prec),
                    )

                    # Same text-only expiry gate as every other extractor: a
                    # source-wide date scan is a display fallback only.
                    if nt.is_expired(text_dates[2] or validity_end, text_dates[3] or ve_prec):
                        continue

                    cities = nt.detect_cities(b, nt.INDIAN_STATE_CITIES.get(state, {}))

                    relevance = nt.compute_gcc_relevance(b, "India")

                    base = {
                        "country_name": "India",
                        "state_name": state,
                        "description": b,
                        "validity_start_date": start,
                        "validity_start_precision": s_prec,
                        "validity_end_date": end,
                        "validity_end_precision": e_prec,
                        "gcc_relevance": relevance,
                    }

                    if cities:
                        for city_name, tier in cities:
                            rows_by_state[state].append({
                                **base, "section_name": section, "scope_type": "City",
                                "city_name": city_name, "tier_scope": tier,
                            })
                    else:
                        rows_by_state[state].append({
                            **base, "section_name": section, "scope_type": scope_type,
                            "city_name": "", "tier_scope": tier_scope,
                        })

    for state in rows_by_state:
        rows_by_state[state] = nt._fuzzy_dedupe_rows(rows_by_state[state])

    return rows_by_state


def output_paths_global(paraphrased=True):

    today = date.today().isoformat()
    variant = "paraphrased" if paraphrased else "raw"

    old_dir = old_template_dir_for("Global")
    os.makedirs(old_dir, exist_ok=True)

    return (
        os.path.join(old_dir, f"gcc_incentives_global_{variant}_{today}.xlsx"),
        os.path.join(old_dir, f"gcc_sources_{variant}_{today}.xlsx"),
    )


def run_global_target(args):

    countries_filter = None

    if getattr(args, "countries", None):
        countries_filter = [c.strip() for c in args.countries.split(",") if c.strip()]

    label = f"Global ({len(countries_filter)} countries, no India)" if countries_filter else "Global (Global sheet + India)"

    print(f"\n{'=' * 70}\nTarget: {label}\n{'=' * 70}\n")

    sources_file = os.path.join(SOURCES_DIR, "sources.xlsx")

    if not os.path.exists(sources_file):
        print(f"✗ {sources_file} not found. Skipping Global.")
        return

    print(f"Loading sources from {sources_file}\n")

    srcs = SourcesLoader.load(sources_file)

    if countries_filter:

        # A targeted subset run: skip India entirely (not asked for), and
        # restrict the country list to exactly what was requested rather
        # than the full ~330-country template list.
        india_states = []
        all_countries = countries_filter

        unconfigured = [c for c in all_countries if c not in srcs["country_pdfs"] and c not in srcs["country_html"]]
        if unconfigured:
            print(f"  ⚠ no PDF/HTML sources configured for: {', '.join(unconfigured)} "
                  "(will produce 0 rows unless already present in sources.xlsx)")

    else:

        india_states = _ordered_keys(srcs["india_pdfs"], srcs["india_html"], srcs["india_validity"])

        countries_from_sources = _ordered_keys(
            srcs["country_pdfs"], srcs["country_html"], srcs["country_validity"]
        )

        countries_from_template = TemplateLoader.load_countries(TEMPLATE_FILE)

        all_countries = _merge_lists(countries_from_template, countries_from_sources)

    print(f"  → India states: {len(india_states)}\n  → Countries: {len(all_countries)}\n")

    os.makedirs(INDIA_SOURCES_DIR, exist_ok=True)
    os.makedirs(GLOBAL_SOURCES_DIR, exist_ok=True)

    india_pdfs = srcs["india_pdfs"] if india_states else {}

    country_pdfs = srcs["country_pdfs"]

    if countries_filter:

        # Only fetch / freshness check / refresh the PDFs of the requested
        # countries. This used to pass every configured country's PDFs
        # through regardless of --countries: a China only run first spent
        # about 5 minutes checking and re-downloading ~40 other countries'
        # PDFs (several timing out) before reaching China, and silently
        # replaced other countries' cached source files as a side effect.
        country_pdfs = {c: p for c, p in country_pdfs.items() if c in set(all_countries)}

    india_available, country_available = ensure_pdf_sources_global(
        india_pdfs, country_pdfs, force=args.refresh,
    )

    translator = None

    if args.translate:
        tr = ArgosTranslator()
        if tr.is_available():
            translator = tr
        else:
            print("⚠ argostranslate not available; non-English content will not be translated.")

    state_data = {}
    validity_by_state = {}

    if india_states:

        print(f"\nStep 2: extracting India ({EXTRACT_WORKERS} workers)\n")

        state_data = _build_jurisdiction_data(
            india_available, srcs["india_html"], india_states,
            INDIA_SOURCES_DIR, translator, "India",
        )
        validity_by_state = {s: srcs["india_validity"].get(s, "") for s in india_states}

    print(f"\nStep 3: extracting Global countries ({EXTRACT_WORKERS} workers)\n")

    country_data = _build_jurisdiction_data(
        country_available, srcs["country_html"], all_countries,
        GLOBAL_SOURCES_DIR, translator, "Global",
    )
    validity_by_country = {c: srcs["country_validity"].get(c, "") for c in all_countries}

    # Check against the full country universe (not just the countries in
    # this run) since a cross-country-comparison document discovered under
    # one country's crawl can mention any of them -- confirmed case: Spain's
    # own incentives page links to a PDF comparing it against Germany,
    # France, Poland, Hungary and the Netherlands, none of which need to be
    # part of this particular run for their names to show up in the text.
    known_countries = set(_merge_lists(TemplateLoader.load_countries(TEMPLATE_FILE), all_countries))

    country_data = filter_cross_country_bleed(country_data, known_countries)

    if args.paraphrase:

        para = OllamaParaphraser(model=args.model, min_len=args.min_len)

        if not para.is_available():
            print(
                "⚠ Ollama not reachable/model not pulled; skipping paraphrasing. "
                f"Run: ollama pull {args.model}"
            )
        else:
            print("\nParaphrasing via Ollama\n")
            if state_data:
                apply_transform(state_data, para.paraphrase_many, "states")
            apply_transform(country_data, para.paraphrase_many, "countries")

    if getattr(args, "no_old_template", False):

        print("\n--no-old-template: skipping the wide/audit workbook, building "
              "new_template directly from in-memory data\n")

        filtered_srcs = dict(srcs)
        filtered_srcs["india_pdfs"] = {}
        filtered_srcs["india_html"] = {}
        filtered_srcs["country_pdfs"] = {
            c: p for c, p in srcs["country_pdfs"].items() if c in country_data
        }
        filtered_srcs["country_html"] = {
            c: h for c, h in srcs["country_html"].items() if c in country_data
        }

        source_dates_index = nt.build_source_dates_index(filtered_srcs)

        rows_by_country = extract_global_rows_from_data(
            country_data, validity_by_country, source_dates_index,
            known_countries=known_countries,
        )

        all_rows = []
        for country in rows_by_country:
            all_rows.extend(sorted(rows_by_country[country], key=lambda r: -(r.get("gcc_relevance") or 0)))

        signal_rows = [r for r in all_rows if int(r.get("gcc_relevance") or 0) >= 1]
        other_rows = [r for r in all_rows if int(r.get("gcc_relevance") or 0) == 0]

        new_dir = new_template_dir_for("Global")
        os.makedirs(new_dir, exist_ok=True)

        today = date.today().isoformat()
        suffix = "_subset" if countries_filter else ""
        flat_path = os.path.join(new_dir, f"new_template_GCC_policy{suffix}_{today}.xlsx")

        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        nt._write_flat_sheet(wb, "Global", signal_rows)
        nt._write_flat_sheet(wb, "Other - Global", other_rows)
        wb.save(flat_path)

        print(f"\n✓ Wrote {flat_path}  (Global: {len(signal_rows)} rows, Other - Global: {len(other_rows)} rows)")
        return

    print("\nWriting workbooks\n")

    out_path, sources_path = output_paths_global(paraphrased=args.paraphrase)

    baseline_prefix = (
        "gcc_incentives_global_paraphrased_" if args.paraphrase else "gcc_incentives_global_raw_"
    )

    baseline_path = PreviousWorkbookReader.find_latest(
        old_template_dir_for("Global"), baseline_prefix, exclude_path=out_path,
    )

    baseline = PreviousWorkbookReader.load(baseline_path)

    if baseline:
        print(f"  ⟂ comparing against baseline: {os.path.basename(baseline_path)}")
    else:
        print("  ⟂ no prior Global workbook found; all content rendered black")

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

    rows = build_global_sources_rows(srcs, india_available, country_available)

    SourcesExcelExporter.export(rows, sources_path)

    print(f"  ✓ Wrote {sources_path}")

    print("\nBuilding new_template flat workbook for Global + India\n")

    try:
        nt.run()
    except Exception as exc:
        print(f"⚠ new_template.run() failed: {exc}. You can run 'python3 new_template.py' manually.")


def run_india_target(args):
    """India's GCC policy states on their own, without the ~330 country Global
    sheet. --global already includes India (same states, same sources), so
    this exists only to refresh India without paying for a full Global run.
    Outputs go to their own old_template/India and new_template/India folders
    so they can never be mistaken for, or overwrite, a Global run."""

    print(f"\n{'=' * 70}\nTarget: India (8 GCC policy states, no Global countries)\n{'=' * 70}\n")

    sources_file = os.path.join(SOURCES_DIR, "sources.xlsx")

    if not os.path.exists(sources_file):
        print(f"✗ {sources_file} not found. Skipping India.")
        return

    print(f"Loading sources from {sources_file}\n")

    srcs = SourcesLoader.load(sources_file)

    india_states = _ordered_keys(srcs["india_pdfs"], srcs["india_html"], srcs["india_validity"])

    print(f"  → India states: {len(india_states)}\n")

    os.makedirs(INDIA_SOURCES_DIR, exist_ok=True)

    india_available, _ = ensure_pdf_sources_global(srcs["india_pdfs"], {}, force=args.refresh)

    translator = None

    if args.translate:
        tr = ArgosTranslator()
        if tr.is_available():
            translator = tr

    print(f"\nExtracting India ({EXTRACT_WORKERS} workers)\n")

    state_data = _build_jurisdiction_data(
        india_available, srcs["india_html"], india_states,
        INDIA_SOURCES_DIR, translator, "India",
    )
    validity_by_state = {s: srcs["india_validity"].get(s, "") for s in india_states}

    if args.paraphrase:

        para = OllamaParaphraser(model=args.model, min_len=args.min_len)

        if para.is_available():
            print("\nParaphrasing via Ollama\n")
            apply_transform(state_data, para.paraphrase_many, "states")
        else:
            print("⚠ Ollama not reachable/model not pulled; skipping paraphrasing.")

    india_only_srcs = dict(srcs)
    india_only_srcs["country_pdfs"] = {}
    india_only_srcs["country_html"] = {}
    india_only_srcs["country_validity"] = {}

    source_dates_index = nt.build_source_dates_index(india_only_srcs)

    rows_by_state = extract_india_rows_from_data(state_data, validity_by_state, source_dates_index)

    # Tamil Nadu is a local file (no URL); same fallback new_template.run()
    # has if its bucketed content produced nothing.
    if not rows_by_state.get("Tamil Nadu"):
        tn_rows = nt.extract_tamil_nadu_from_pdf()
        if tn_rows:
            rows_by_state["Tamil Nadu"] = tn_rows

    all_rows = []
    for state, rows in rows_by_state.items():
        all_rows.extend(sorted(rows, key=lambda r: -int(r.get("gcc_relevance") or 0)))

    today = date.today().isoformat()

    new_dir = new_template_dir_for("India")
    os.makedirs(new_dir, exist_ok=True)

    flat_path = os.path.join(new_dir, f"new_template_india_{today}.xlsx")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    nt._write_flat_sheet(wb, "India", all_rows)
    wb.save(flat_path)

    per_state = {s: len(r) for s, r in rows_by_state.items()}

    print(f"\n✓ Wrote {flat_path}  (India: {len(all_rows)} rows across "
          f"{sum(1 for n in per_state.values() if n)} of {len(per_state)} states)")
    print("  rows per state:", per_state)

    if getattr(args, "no_old_template", False):
        print("  (--no-old-template: wide workbook and gcc_sources file skipped)")
        return

    old_dir = old_template_dir_for("India")
    os.makedirs(old_dir, exist_ok=True)

    variant = "paraphrased" if args.paraphrase else "raw"

    wide_path = os.path.join(old_dir, f"gcc_incentives_india_{variant}_{today}.xlsx")

    baseline_path = PreviousWorkbookReader.find_latest(
        old_dir, f"gcc_incentives_india_{variant}_", exclude_path=wide_path,
    )

    PolicyExcelExporter.export_global(
        state_data=state_data,
        validity_by_state=validity_by_state,
        country_data={},
        validity_by_country={},
        placeholder_data={},
        placeholder_validity={},
        output_path=wide_path,
        baseline=PreviousWorkbookReader.load(baseline_path),
    )

    print(f"✓ Wrote {wide_path}")

    sources_path = os.path.join(old_dir, f"gcc_sources_india_{variant}_{today}.xlsx")

    rows = build_global_sources_rows(india_only_srcs, india_available, {})

    SourcesExcelExporter.export(rows, sources_path)

    print(f"✓ Wrote {sources_path}  ({len(rows)} source rows)")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

TARGETS = ["United States", "United Kingdom", "Germany", "France", "Japan", "Global"]

# One boolean flag per target. --global builds India's 8 states together
# with the ~330-country Global sheet; --india builds just those 8 states.
# India is deliberately NOT in TARGETS (the default run-everything list)
# because Global already covers it.
FLAG_TO_TARGET = {
    "usa": "United States",
    "uk": "United Kingdom",
    "germany": "Germany",
    "france": "France",
    "japan": "Japan",
    "global": "Global",
    "india": "India",
}


def run_target(target, args):

    if target == "Global":
        run_global_target(args)
    elif target == "India":
        run_india_target(args)
    else:
        run_state_level_target(target, args)


def main():

    parser = argparse.ArgumentParser(
        description=(
            "Build GCC incentive workbooks. Default: run every target "
            f"({', '.join(TARGETS)}), one after another. Pass one or more "
            "per-country flags to run only those."
        )
    )

    for flag, target in FLAG_TO_TARGET.items():
        parser.add_argument(
            f"--{flag}", action="store_true",
            help=f"Run only the {target} target. Combine flags to run more than one.",
        )

    parser.add_argument("--refresh", action="store_true",
                         help="Force re-download of all PDFs, ignoring local cache.")
    parser.add_argument("--nopara", "--no-paraphrase", dest="paraphrase",
                         action="store_false", default=True,
                         help="Skip the Ollama paraphrasing step.")
    parser.add_argument("--notranslate", "--no-translate", dest="translate",
                         action="store_false", default=True,
                         help="Skip the argos-translate step.")
    parser.add_argument("--model", default=OllamaParaphraser.DEFAULT_MODEL,
                         help=f"Ollama model name for paraphrasing (default: {OllamaParaphraser.DEFAULT_MODEL})")
    parser.add_argument("--min-len", type=int, default=0,
                         help="Minimum bullet length before paraphrasing kicks in (default: 0, "
                              "i.e. paraphrase everything).")
    parser.add_argument("--no-old-template", action="store_true",
                         help="Skip writing the old_template (wide/audit) workbook and, for "
                              "state-level targets, its gcc_sources file. Only the flat "
                              "new_template workbook is produced.")
    parser.add_argument("--max-minutes", type=float, default=None,
                         help="Stop cleanly after about this many minutes, at the next "
                              "paraphrase chunk, with progress saved (exit code 3). Run the "
                              "same command again to continue.")
    parser.add_argument("--resume", action="store_true",
                         help="Keep the fetched and translated text of an interrupted run "
                              "(output/.resume, deleted when the run completes) so running the "
                              "same command again does not fetch everything again. Saved text "
                              "older than 24 hours is ignored.")
    parser.add_argument("--countries", default=None,
                         help="Global target only: comma-separated list of specific countries "
                              "to (re)build, e.g. --global --countries \"Brazil,Spain,Italy\". "
                              "Skips India and every other country. Default: all ~330 countries "
                              "plus India.")

    args = parser.parse_args()

    global RESUME
    RESUME = args.resume

    if args.max_minutes is not None:
        OllamaParaphraser.DEADLINE = time.time() + args.max_minutes * 60

    selected = [target for flag, target in FLAG_TO_TARGET.items() if getattr(args, flag)]

    targets = selected or TARGETS

    # --global already builds India (same states, same sources); running both
    # would just do India twice. A --countries subset run skips India, so the
    # two don't overlap there.
    if "India" in targets and "Global" in targets and not args.countries:
        print("ℹ --global already includes India; skipping the separate India run.")
        targets = [t for t in targets if t != "India"]

    try:

        for target in targets:
            run_target(target, args)
            _resume_clear()

    except ParaphraseDeadline as exc:

        print(
            f"\n{'=' * 70}\n⏸ Stopped: {exc}. Paraphrase progress is saved"
            f"{' and the fetched text is kept (--resume)' if RESUME else ''}. "
            f"Run the same command again to continue.\n{'=' * 70}"
        )

        sys.exit(3)

    print(f"\n{'=' * 70}\nDone: {', '.join(targets)}\n{'=' * 70}")


if __name__ == "__main__":
    main()
