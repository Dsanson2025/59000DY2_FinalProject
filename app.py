"""
Louisiana Sheriff Sale Scraper — with Market Value & Profit Estimates
======================================================================
Scrapes sheriff sale / foreclosure listings for major Louisiana metro
parishes and outputs an Excel file with:
  - Property Type classification (Commercial / Land / Multi-Family / Residential)
  - Avg $/sqft for the property zip code (from Redfin public market data, no key needed)
  - Square footage of the specific property (optional, requires free Rentcast API key)
  - Estimated Market Value  = sqft x avg $/sqft   (falls back to zip-level comps
                                                     when sqft isn't available)
  - Profit Estimate         = Market Value - Writ/Sale Price

SETUP:
    pip install requests beautifulsoup4 openpyxl lxml

    Optional (for square footage data — unlocks the sqft-based Market Value +
    Profit columns; without it, Market Value falls back to the zip's median
    comp value from Redfin/Census):
      1. Sign up free at https://app.rentcast.io  (50 requests/month on free tier)
      2. Get your key at https://app.rentcast.io/app/api-keys
      3. Either set an env var:  set RENTCAST_API_KEY=your_key_here  (Windows)
                                 export RENTCAST_API_KEY=your_key_here  (Mac/Linux)
         Or paste it directly into the RENTCAST_API_KEY variable below.

RUN:
    python sheriff_sale_scraper.py

OUTPUT:
    sheriff_sales_commercial.xlsx  (in the same folder)
"""
import re
import os
import csv
import gzip
import io
import time
import logging
from dataclasses import dataclass
from typing import Optional
from urllib.parse import urljoin
from urllib.parse import quote_plus

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(name)s | %(message)s")
log = logging.getLogger("sheriff_scraper")

# ── Market data caches (populated lazily on first use) ────────────────────────
_zip_data_cache: dict[str, dict] = {}   # zip → {psf, median_sale_price, median_list_price}
_census_cache: dict[str, Optional[float]] = {}  # zip → Census median home value
_redfin_loaded = False
_census_loaded = False

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
}
REQUEST_TIMEOUT = 20
REQUEST_DELAY_SECONDS = 1.5  # be polite, avoid hammering small govt sites

# ── Rentcast (optional square-footage lookup) ─────────────────────────────────
# Paste your key here, or set the RENTCAST_API_KEY environment variable instead
# (the env var takes precedence if both are set... actually: env var wins only
# if this literal is left blank, since that's the safer default for sharing
# this script with a key already filled in).
RENTCAST_API_KEY = os.environ.get("RENTCAST_API_KEY", "").strip() or ""
RENTCAST_PROPERTY_URL = "https://api.rentcast.io/v1/properties"
# Free tier is 50 requests/month - cap a bit under that so one run never
# silently exhausts the whole month's budget by itself.
MAX_RENTCAST_CALLS_PER_RUN = 45

_sqft_cache: dict[str, Optional[int]] = {}  # normalized address → sqft (or None)
_rentcast_calls_made = 0
_rentcast_budget_warned = False

COMMERCIAL_KEYWORDS = [
    "commercial", "retail", "office", "warehouse", "industrial",
    "business", "storefront", "strip mall", "plaza", "restaurant",
    "gas station", "c-store", "mixed use", "zoned c-", "zoned b-",
    "shopping center", "professional building", "medical building",
    "auto repair", "salon", "laundromat", "motel", "hotel",
    "church", "daycare", "academy", "lounge", "bar &", "tavern",
]
LAND_KEYWORDS = [
    "vacant lot", "vacant land", "unimproved land", "unimproved lot",
    "vacant tract", "acreage", "acres", "acre tract", "raw land",
    "undeveloped", "buildable lot", "lot only", "land only",
    "no improvements", "tract of land", "parcel of land",
]
# signal a commercial property even when the address text doesn't say so -
# residential foreclosures are filed against individual people.
ENTITY_SUFFIX_PATTERN = re.compile(
    r"\b(LLC|L\.L\.C\.|INC|L\.P\.|LP|CORP|CORPORATION|COMPANY|CO\.|ENTERPRISES|"
    r"PARTNERS|HOLDINGS|GROUP|ASSOCIATES|VENTURES|PROPERTIES|REALTY|"
    r"INVESTMENTS|ACADEMY|MINISTRIES|CHURCH)\b", re.IGNORECASE
)
RESIDENTIAL_KEYWORDS = [
    "single family", "single-family", "sfr", "residence", "residential",
    "townhome", "townhouse", "condo", "condominium", "subdivision lot",
    "bedroom", "bath home", "house",
]
MULTI_FAMILY_KEYWORDS = [
    "duplex", "triplex", "fourplex", "quadplex", "multi-family",
    "multi family", "apartment", "apartments", "units", "2-family",
    "2 family", "3-family", "3 family", "4-family", "4 family",
]


@dataclass
class Listing:
    parish: str
    address: str
    price: Optional[float]
    sale_date: str
    description: str
    source_url: str
    property_type: str = "Unknown"  # "Commercial", "Land", "Multi-Family", "Residential", "Unknown"
    status: str = "OK"  # OK, NEEDS_FIX, NO_DATA
    case_number: Optional[str] = None  # populated by parish-specific parsers, used only for de-duping

    @property
    def likely_commercial(self) -> bool:
        return self.property_type == "Commercial"


def _extract_zip(address: str) -> Optional[str]:
    """Pull 5-digit zip code out of an address string."""
    m = re.search(r"\b(7[01]\d{3})\b", address)  # Louisiana zips: 70000-71499
    return m.group(1) if m else None


def _load_redfin_zip_data():
    """
    Stream Redfin's public zip-level market tracker without loading the
    entire 1.5GB file into memory. Filters Louisiana zips on the fly.
    Column names in this file are UPPERCASE (REGION, MEDIAN_PPSF, etc).
    """
    global _redfin_loaded
    if _redfin_loaded:
        return
    _redfin_loaded = True

    url = ("https://redfin-public-data.s3.us-west-2.amazonaws.com/"
           "redfin_market_tracker/zip_code_market_tracker.tsv000.gz")
    log.info("Streaming Redfin zip market data (filtering Louisiana rows only) ...")
    try:
        with requests.get(url, stream=True, timeout=120,
                          headers={**HEADERS, "Referer": "https://www.redfin.com/"}) as resp:
            resp.raise_for_status()
            # Stream-decompress without buffering the full 1.5GB file
            resp.raw.decode_content = True
            with gzip.GzipFile(fileobj=resp.raw) as gz:
                reader = csv.DictReader(
                    io.TextIOWrapper(gz, encoding="utf-8", errors="replace"),
                    delimiter="\t"
                )
                best: dict[str, dict] = {}
                for row in reader:
                    region = row.get("REGION", "")
                    m = re.search(r"\b(7[01]\d{3})\b", region)
                    if not m:
                        continue
                    z = m.group(1)
                    period = row.get("PERIOD_END", "")
                    if z not in best or period > best[z].get("PERIOD_END", ""):
                        best[z] = row

        def _f(row, key):
            try:
                v = row.get(key, "").strip()
                return float(v) if v else None
            except ValueError:
                return None

        for z, row in best.items():
            _zip_data_cache[z] = {
                "psf":               _f(row, "MEDIAN_PPSF"),
                "median_sale_price": _f(row, "MEDIAN_SALE_PRICE"),
                "median_list_price": _f(row, "MEDIAN_LIST_PRICE"),
            }
        log.info("Redfin data loaded: %d Louisiana zip codes", len(_zip_data_cache))
    except Exception as e:
        log.warning("Could not load Redfin market data: %s", e)


def _load_census_zip_data():
    """
    Pull Census Bureau ACS 5-year median home value (B25077_001E) for all
    Louisiana zip code tabulation areas. Free, no API key required.
    Only runs once per session.
    """
    global _census_loaded
    if _census_loaded:
        return
    _census_loaded = True

    # ACS 5-year estimates - use most recent stable year
    url = ("https://api.census.gov/data/2022/acs/acs5"
           "?get=B25077_001E&for=zip+code+tabulation+area:7*&in=state:22")
    log.info("Downloading Census Bureau median home values for Louisiana ...")
    try:
        resp = requests.get(url, timeout=30, headers=HEADERS)
        resp.raise_for_status()
        data = resp.json()
        # data[0] is the header row: ['B25077_001E', 'state', 'zip code tabulation area']
        for row in data[1:]:
            val_str, state, zcta = row
            try:
                val = float(val_str)
                if val > 0:  # Census returns -666666666 for suppressed data
                    _census_cache[zcta] = val
                else:
                    _census_cache[zcta] = None
            except (ValueError, TypeError):
                _census_cache[zcta] = None
        log.info("Census data loaded: %d Louisiana zip codes", len(_census_cache))
    except Exception as e:
        log.warning("Could not load Census median home value data: %s", e)


def get_zip_market_data(address: str) -> dict:
    """
    Returns a dict with market data for the zip code in address:
      psf               – Redfin median sale $/sqft
      median_sale_price – Redfin median sale price
      median_list_price – Redfin median list price
      census_value      – Census Bureau ACS median home value (5-yr estimate)
      market_value      – fallback market value estimate when a property-specific
                           sqft x psf calculation isn't available (sale > census)
    All values are floats or None.
    """
    _load_redfin_zip_data()
    _load_census_zip_data()
    z = _extract_zip(address)
    if not z:
        return {}
    redfin = _zip_data_cache.get(z, {})
    census_val = _census_cache.get(z)
    market_value = redfin.get("median_sale_price") or census_val
    return {
        "zip":                z,
        "psf":                redfin.get("psf"),
        "median_sale_price":  redfin.get("median_sale_price"),
        "median_list_price":  redfin.get("median_list_price"),
        "census_value":       census_val,
        "market_value":       market_value,
    }


def get_property_sqft(address: str) -> Optional[int]:
    """
    Look up square footage for a specific property via the Rentcast API
    (free tier: 50 requests/month — see RENTCAST_API_KEY setup notes at the
    top of this file).

    Returns None (and makes no network call) when:
      - no API key is configured, or
      - this run has already hit MAX_RENTCAST_CALLS_PER_RUN, or
      - the lookup fails or returns no match.

    Results are cached by normalized address so the same property is never
    looked up twice in a single run.
    """
    global _rentcast_calls_made, _rentcast_budget_warned

    if not RENTCAST_API_KEY:
        return None

    key = address.strip().lower()
    if key in _sqft_cache:
        return _sqft_cache[key]

    if _rentcast_calls_made >= MAX_RENTCAST_CALLS_PER_RUN:
        if not _rentcast_budget_warned:
            log.warning(
                "Rentcast call budget (%d) reached for this run - remaining "
                "listings will fall back to zip-level comps for Market Value.",
                MAX_RENTCAST_CALLS_PER_RUN,
            )
            _rentcast_budget_warned = True
        return None

    try:
        resp = requests.get(
            RENTCAST_PROPERTY_URL,
            params={"address": address},
            headers={"X-Api-Key": RENTCAST_API_KEY, "Accept": "application/json"},
            timeout=REQUEST_TIMEOUT,
        )
        _rentcast_calls_made += 1
        time.sleep(REQUEST_DELAY_SECONDS)

        if resp.status_code == 429:
            log.warning("Rentcast rate limit hit (HTTP 429) for %s", address)
            _sqft_cache[key] = None
            return None
        resp.raise_for_status()

        data = resp.json()
        record = None
        if isinstance(data, list) and data:
            record = data[0]
        elif isinstance(data, dict):
            record = data

        sqft = None
        if record:
            raw_sqft = record.get("squareFootage")
            if raw_sqft:
                try:
                    sqft = int(raw_sqft)
                except (ValueError, TypeError):
                    sqft = None

        _sqft_cache[key] = sqft
        return sqft
    except Exception as e:
        log.warning("Rentcast lookup failed for %s: %s", address, e)
        _sqft_cache[key] = None
        return None


def parse_price(text: str) -> Optional[float]:
    if not text:
        return None
    match = re.search(r"\$\s?([\d,]+(?:\.\d{2})?)", text)
    if not match:
        return None
    try:
        return float(match.group(1).replace(",", ""))
    except ValueError:
        return None


def classify_property_type(text: str) -> str:
    t = text.lower()
    if any(k in t for k in COMMERCIAL_KEYWORDS):
        return "Commercial"
    if ENTITY_SUFFIX_PATTERN.search(text):
        return "Commercial"
    if any(k in t for k in MULTI_FAMILY_KEYWORDS):
        return "Multi-Family"
    if any(k in t for k in LAND_KEYWORDS):
        return "Land"
    if any(k in t for k in RESIDENTIAL_KEYWORDS):
        return "Residential"
    return "Unknown"


def fetch(url: str) -> Optional[BeautifulSoup]:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        time.sleep(REQUEST_DELAY_SECONDS)
        return BeautifulSoup(resp.text, "lxml")
    except requests.RequestException as e:
        log.warning("Fetch failed for %s: %s", url, e)
        return None


# ---------------------------------------------------------------------------
# PARISH SCRAPERS
# Each function returns a list[Listing]. Mark status="NEEDS_FIX" when the
# page structure can't be confirmed yet, but still try a best-effort parse.
# ---------------------------------------------------------------------------

def _parse_civilview_table(soup: BeautifulSoup, parish: str, url: str) -> list[Listing]:
    """Parse the CivilView sales results table into Listings.

    The page has multiple <table> elements (banner, nav, data).
    Find the right one by looking for known column headers.
    """
    listings: list[Listing] = []

    # Find the sales data table specifically — not the first table on
    # the page (which is the OPSO banner/welcome table)
    sales_table = None
    for tbl in soup.find_all("table"):
        header_text = tbl.get_text(" ", strip=True).lower()
        if "case #" in header_text or "sort order" in header_text or "writ amount" in header_text:
            sales_table = tbl
            break

    if sales_table is None:
        return listings

    rows = sales_table.find_all("tr")
    for row in rows[1:]:  # skip header row
        cells = row.find_all("td")
        if len(cells) < 6:
            continue

        # Confirmed column layout from live page inspection:
        # 0=Sort Order, 1=Case#, 2=Sales Date, 3=Property Status,
        # 4=Case Title, 5=Address/Description, 6=Picture,
        # 7=Attorney, 8=Writ Amount, 9=Terms, 10=Details
        case_number  = cells[1].get_text(" ", strip=True)
        sale_date    = cells[2].get_text(" ", strip=True)
        case_title   = cells[4].get_text(" ", strip=True)
        address      = cells[5].get_text(" ", strip=True)
        writ_text    = cells[8].get_text(" ", strip=True) if len(cells) > 8 else ""
        full_text    = f"{case_title} {address}"
        price        = parse_price(writ_text) or parse_price(full_text)

        listing = Listing(parish, address or "See case details", price, sale_date,
                           case_title, url,
                           property_type=classify_property_type(full_text),
                           status="OK", case_number=case_number)
        listings.append(listing)

    return listings


def scrape_civilview(parish: str, county_id: int) -> list[Listing]:
    """
    VERIFIED WORKING. CivilView SalesWeb used by Orleans (countyId=28)
    and Ascension (countyId=55).

    The default GET to ?countyId=28 already returns the next upcoming
    sale date's full listing with no form submission needed. For future
    dates, the URL parameter is SalesDate=MM/DD/YYYY (confirmed from
    live page inspection). We loop all upcoming dates from the dropdown.
    """
    import datetime

    base_url = f"https://salesweb.civilview.com/Sales/SalesSearch?countyId={county_id}"
    session = requests.Session()
    session.headers.update(HEADERS)

    # Step 1: fetch the default page — this always returns the next
    # upcoming sale date's results and gives us the full date dropdown.
    try:
        resp = session.get(base_url, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
    except requests.RequestException as e:
        log.warning("Initial fetch failed for %s: %s", parish, e)
        return [Listing(parish, "", None, "", f"Fetch failed: {e}", base_url, status="NEEDS_FIX")]

    soup = BeautifulSoup(resp.text, "lxml")

    # Step 2: find the date dropdown to get all upcoming dates.
    # Option values are plain MM/DD/YYYY strings (not numeric IDs).
    date_select = None
    for sel in soup.find_all("select"):
        opts = [o.get("value", "") for o in sel.find_all("option")]
        if any(re.match(r"\d{1,2}/\d{1,2}/\d{4}$", v) for v in opts if v):
            date_select = sel
            break

    today = datetime.date.today()
    upcoming_dates: list[str] = []
    if date_select:
        for opt in date_select.find_all("option"):
            val = (opt.get("value") or "").strip()
            m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})$", val)
            if not m:
                continue
            try:
                d = datetime.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
            except ValueError:
                continue
            if d >= today:
                upcoming_dates.append(val)
    upcoming_dates.sort()

    if not upcoming_dates:
        # No dropdown found — just parse whatever the default page returned
        log.info("[%s] No date dropdown found, parsing default page only.", parish)
        listings = _parse_civilview_table(soup, parish, base_url)
        if not listings:
            return [Listing(parish, "", None, "", "Ran OK - 0 listings on default page",
                             base_url, status="NO_DATA")]
        return listings

    log.info("[%s] found %d upcoming sale date(s)", parish, len(upcoming_dates))
    all_listings: list[Listing] = []
    seen_case_numbers: set[str] = set()

    for date_val in upcoming_dates:
        # Build URL directly — no form submission needed, server accepts
        # SalesDate as a plain GET parameter alongside countyId
        url = (f"https://salesweb.civilview.com/Sales/SalesSearch"
               f"?countyId={county_id}&SalesDate={date_val}")
        try:
            r = session.get(url, timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
        except requests.RequestException as e:
            log.warning("[%s] failed to fetch %s: %s", parish, date_val, e)
            continue
        time.sleep(REQUEST_DELAY_SECONDS)

        page_soup = BeautifulSoup(r.text, "lxml")
        for listing in _parse_civilview_table(page_soup, parish, base_url):
            case_no = listing.case_number
            if case_no and case_no in seen_case_numbers:
                continue
            if case_no:
                seen_case_numbers.add(case_no)
            all_listings.append(listing)

    # Always include the default page results as a safety net in case
    # the SalesDate parameter doesn't work for every date
    for listing in _parse_civilview_table(soup, parish, base_url):
        case_no = listing.case_number
        if case_no and case_no in seen_case_numbers:
            continue
        if case_no:
            seen_case_numbers.add(case_no)
        all_listings.append(listing)

    if not all_listings:
        return [Listing(parish, "", None, "",
                         "Ran OK - 0 listings found across all upcoming dates",
                         base_url, status="NO_DATA")]
    return all_listings


def _parse_jpso_table(soup: BeautifulSoup, parish: str, url: str) -> list[Listing]:
    """Parse one rendered Real Estate Sales List table into Listings."""
    listings: list[Listing] = []
    table = soup.find("table")
    if table is None:
        return listings

    rows = table.find_all("tr")
    for row in rows[1:]:
        cells = row.find_all("td")
        if len(cells) < 5:
            continue

        case_number = cells[0].get_text(" ", strip=True)
        case_style = cells[1].get_text(" ", strip=True)
        sale_date = cells[2].get_text(" ", strip=True)
        address = cells[3].get_text(" ", strip=True)
        writ_amount_text = cells[4].get_text(" ", strip=True)
        full_text = f"{case_style} {address}"
        price = parse_price(writ_amount_text)

        listing = Listing(parish, address or "Address not available", price, sale_date,
                           case_style, url, property_type=classify_property_type(full_text),
                           status="OK", case_number=case_number)
        listings.append(listing)

    return listings


def scrape_jpso(parish: str = "Jefferson") -> list[Listing]:
    """
    VERIFIED WORKING. Jefferson Parish Sheriff's Office real estate sales
    list. The page is a POST-back form: choosing a date in the dropdown
    re-submits the page and swaps the table. Rather than guessing the
    field name, this auto-discovers it from the live form, then loops
    every date in the dropdown that is today or later, de-duping by case
    number (the same case sometimes reappears across multiple weeks while
    pending).
    """
    import datetime

    base_url = "https://eservices2.jpso.com/JudpSale/Home/RealEstate"
    session = requests.Session()
    session.headers.update(HEADERS)

    try:
        resp = session.get(base_url, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
    except requests.RequestException as e:
        log.warning("Initial fetch failed for Jefferson: %s", e)
        return [Listing(parish, "", None, "", f"Fetch failed: {e}", base_url, status="NEEDS_FIX")]

    soup = BeautifulSoup(resp.text, "lxml")

    form = None
    date_select = None
    # find the <select> whose options look like dates, and its parent <form>
    for sel in soup.find_all("select"):
        options = [o.get_text(strip=True) for o in sel.find_all("option")]
        if any(re.match(r"\d{1,2}/\d{1,2}/\d{4}", o) for o in options if o):
            date_select = sel
            form = sel.find_parent("form")
            break

    if date_select is None or form is None:
        log.warning("Could not locate the date dropdown/form on JPSO page - "
                     "site structure may have changed.")
        # fall back to whatever the default page already shows
        listings = _parse_jpso_table(soup, parish, base_url)
        if not listings:
            return [Listing(parish, "", None, "", "Could not find date selector or table",
                             base_url, status="NEEDS_FIX")]
        return listings

    select_name = date_select.get("name")
    method = (form.get("method") or "get").lower()
    action = form.get("action") or base_url
    action_url = urljoin(base_url, action)

    # collect every other input in the form so we don't drop required hidden
    # fields (anti-forgery tokens etc.) when we resubmit it
    base_payload = {}
    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name:
            continue
        input_type = (inp.get("type") or "text").lower()
        if input_type in ("radio", "checkbox"):
            if inp.has_attr("checked"):
                base_payload[name] = inp.get("value", "")
            elif name not in base_payload:
                base_payload.setdefault(name, None)
        else:
            base_payload[name] = inp.get("value", "")
    base_payload = {k: v for k, v in base_payload.items() if v is not None}

    for sel in form.find_all("select"):
        name = sel.get("name")
        if not name or name == select_name:
            continue
        selected_opt = sel.find("option", selected=True) or sel.find("option")
        if selected_opt is not None:
            base_payload[name] = selected_opt.get("value", selected_opt.get_text(strip=True))

    today = datetime.date.today()
    all_dates = []
    for opt in date_select.find_all("option"):
        val = (opt.get("value") or opt.get_text(strip=True) or "").strip()
        m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", val)
        if not m:
            continue
        try:
            d = datetime.date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
        except ValueError:
            continue
        if d >= today:
            all_dates.append((d, val))
    all_dates.sort()

    if not all_dates:
        log.warning("No upcoming dates found in Jefferson dropdown - using default page only.")
        return _parse_jpso_table(soup, parish, base_url)

    log.info("Jefferson: found %d upcoming sale date(s) to scrape", len(all_dates))
    all_listings: list[Listing] = []
    seen_case_numbers: set[str] = set()
    debug_info = []

    for d, date_value in all_dates:
        payload = dict(base_payload)
        payload[select_name] = date_value
        try:
            if method == "post":
                resp = session.post(action_url, data=payload, timeout=REQUEST_TIMEOUT)
            else:
                resp = session.get(action_url, params=payload, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
        except requests.RequestException as e:
            log.warning("Jefferson: failed to fetch date %s: %s", date_value, e)
            continue
        time.sleep(REQUEST_DELAY_SECONDS)

        page_soup = BeautifulSoup(resp.text, "lxml")
        day_listings = _parse_jpso_table(page_soup, parish, base_url)
        if len(debug_info) < 1:
            debug_info.append(f"first attempt url={resp.url} status={resp.status_code} "
                               f"resp_len={len(resp.text)} rows_parsed={len(day_listings)}")

        for listing in day_listings:
            case_number = listing.case_number
            if case_number and case_number in seen_case_numbers:
                continue
            if case_number:
                seen_case_numbers.add(case_number)
            all_listings.append(listing)

    if not all_listings:
        # Safety net: fall back to the plain default-page fetch that was
        # confirmed working (19 results) before date-looping was added.
        log.warning("Jefferson: date-loop returned 0 results - falling back to default page. "
                    "Debug: %s", "; ".join(debug_info) or "no requests attempted")
        fallback_soup = fetch(base_url)
        fallback_listings = _parse_jpso_table(fallback_soup, parish, base_url) if fallback_soup else []
        if fallback_listings:
            for listing in fallback_listings:
                listing.description += " (NOTE: from default page only - full date-range loop returned 0, needs debugging)"
            return fallback_listings
        debug_text = "; ".join(debug_info) or "no date requests succeeded"
        return [Listing(parish, "", None, "",
                         f"Ran OK - 0 sale(s) found across all upcoming dates AND default page fallback also empty. Debug: {debug_text}",
                         base_url, status="NO_DATA")]

    return all_listings


def scrape_generic_sheriff_site(parish: str, url: str,
                                 row_selector: str,
                                 address_selector: str = "",
                                 price_selector: str = "",
                                 date_selector: str = "") -> list[Listing]:
    """
    Generic best-effort scraper for a parish sheriff's office "Sheriff Sales"
    page. Pass in the page URL and a CSS selector that matches each sale
    row/card. Falls back to scanning all text for $ amounts + commercial
    keywords if specific sub-selectors aren't supplied or don't match.
    """
    listings: list[Listing] = []
    soup = fetch(url)
    if soup is None:
        return [Listing(parish, "", None, "", "Fetch failed", url, status="NEEDS_FIX")]

    rows = soup.select(row_selector) if row_selector else []
    if not rows:
        # fallback: treat whole page text as one blob, split on common separators
        text_blob = soup.get_text("\n", strip=True)
        chunks = re.split(r"\n(?=\d+\.\s|SALE NO|Sale #|SHERIFF SALE)", text_blob)
        rows = chunks

    for row in rows:
        text = row.get_text(" ", strip=True) if hasattr(row, "get_text") else row
        if not text:
            continue

        ptype = classify_property_type(text)
        price = parse_price(text)
        address = ""
        sale_date = ""
        if hasattr(row, "select_one"):
            if address_selector:
                tag = row.select_one(address_selector)
                address = tag.get_text(strip=True) if tag else ""
            if date_selector:
                tag = row.select_one(date_selector)
                sale_date = tag.get_text(strip=True) if tag else ""

        status = "OK" if (price and address) else "NEEDS_FIX"
        listings.append(Listing(parish, address or "See description", price, sale_date, text, url,
                                 property_type=ptype, status=status))

    if not listings:
        listings.append(Listing(parish, "", None, "", "No commercial listings matched on this pass",
                                 url, status="NO_DATA"))
    return listings


# VERIFIED parishes: confirmed live against real HTML on 2026-06-21.
VERIFIED_CIVILVIEW_PARISHES = [
    ("Orleans", 28),
    ("Ascension", 55),
]

# NOT YET VERIFIED: I have not been able to confirm a working scrape source
# for these. Each parish runs its own system and several require manual
# lookup (phone calls, PDF legal notices in local papers, or paid platforms
# like Bid4Assets which block scraping). Listed here so you know what's
# still missing rather than getting a silent gap:
UNVERIFIED_PARISHES = [
    "Caddo", "Calcasieu", "Lafayette",
    "Ouachita", "Rapides", "Terrebonne", "Tangipahoa", "Bossier",
]

# East Baton Rouge: confirmed the sales list page exists at
# https://foreclosure.ebrso.org/RealEstateSales/Index but the table loads
# via JavaScript/AJAX after picking a date - can't be read with a plain
# GET request. To fix: open that page in Chrome, open DevTools (F12) ->
# Network tab -> pick a sale date -> look for an XHR/Fetch request
# (probably returns JSON) -> send me that URL and I'll wire it in directly.
EBR_NOTE = ("JS-rendered table - needs the underlying AJAX/JSON endpoint. "
            "See script comments for how to find it via Chrome DevTools.")


def scrape_st_tammany(parish: str = "St. Tammany") -> list[Listing]:
    """
    VERIFIED WORKING (endpoint confirmed via browser DevTools 2026-06-21).
    St. Tammany's results table is loaded by client-side JS calling:
        GET /Home/RealEstateMovablesSearchResultsTable
            ?SelectedSaleCategory=<id>&SelectedSaleDate=<id>&SelectedSaleStatus=<id>
    Both SelectedSaleCategory and SelectedSaleDate are internal numeric
    database IDs, not literal text/dates - the dropdown's visible text
    (e.g. "Real Estate", "Wednesday, August 26, 2026") is just a label.
    This reads those real id->label mappings directly off the live
    homepage <select> elements (no guessing), then loops every sale date
    today-or-later, requesting "Real Estate" category / "All Sales" status.
    """
    import datetime

    base_url = "https://public.stpso.com/Sheriff.PublicSite/"
    results_endpoint = "https://public.stpso.com/Sheriff.PublicSite/Home/RealEstateMovablesSearchResultsTable"
    fallback_note = (
        "Could not locate the category/date/status dropdowns on the "
        "homepage - site structure may have changed since this was last "
        "verified. Re-check via Chrome DevTools (Network tab -> Fetch/XHR)."
    )

    session = requests.Session()
    session.headers.update(HEADERS)

    try:
        resp = session.get(base_url, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
    except requests.RequestException as e:
        log.warning("Initial fetch failed for St. Tammany: %s", e)
        return [Listing(parish, "", None, "", f"Fetch failed: {e}", base_url, status="NEEDS_FIX")]

    soup = BeautifulSoup(resp.text, "lxml")
    selects = soup.find_all("select")

    def find_select_by_options(*label_substrings):
        for sel in selects:
            option_labels = [o.get_text(strip=True).lower() for o in sel.find_all("option")]
            if any(any(sub in label for label in option_labels) for sub in label_substrings):
                return sel
        return None

    category_select = find_select_by_options("real estate", "movables")
    # NOTE: don't match on "all sales" here - both the category and status
    # dropdowns share that as their first option, which would make this
    # accidentally grab the same <select> as category_select above.
    status_select = find_select_by_options("active", "postponed", "on hold")

    date_select = None
    # the date select is whichever remaining one has the most options
    # (decades of weekly sale dates) and isn't the category/status one
    excluded_ids = {id(category_select), id(status_select)}
    candidate_selects = [s for s in selects if id(s) not in excluded_ids]
    if candidate_selects:
        date_select = max(candidate_selects, key=lambda s: len(s.find_all("option")))

    if category_select is None or date_select is None:
        return [Listing(parish, "", None, "", fallback_note, base_url, status="NEEDS_FIX")]

    def option_value(opt):
        return opt.get("value", opt.get_text(strip=True))

    category_id = None
    for opt in category_select.find_all("option"):
        if "real estate" in opt.get_text(strip=True).lower():
            category_id = option_value(opt)
            break

    status_id = "0"  # "All Sales" - confirmed working value
    if status_select is not None:
        for opt in status_select.find_all("option"):
            if "all sales" in opt.get_text(strip=True).lower():
                status_id = option_value(opt)
                break

    if category_id is None:
        return [Listing(parish, "", None, "", fallback_note, base_url, status="NEEDS_FIX")]

    # build {date: option_id} for every date option that parses and is today or later
    today = datetime.date.today()
    month_names = ("january", "february", "march", "april", "may", "june", "july",
                   "august", "september", "october", "november", "december")
    date_id_pairs = []
    for opt in date_select.find_all("option"):
        label = opt.get_text(strip=True)
        val = option_value(opt)
        if not val or not label:
            continue
        label_low = label.lower()
        if not any(m in label_low for m in month_names):
            continue  # skip the "< Select Sale Date >" placeholder etc.
        m = re.search(r"([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})", label)
        if not m:
            continue
        try:
            d = datetime.datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%B %d %Y").date()
        except ValueError:
            continue
        if d >= today:
            date_id_pairs.append((d, val, label))
    date_id_pairs.sort()

    if not date_id_pairs:
        return [Listing(parish, "", None, "", "Ran OK - no upcoming sale dates found in dropdown",
                         base_url, status="NO_DATA")]

    log.info("[%s] found %d upcoming sale date(s) to scrape", parish, len(date_id_pairs))
    all_listings: list[Listing] = []
    seen_rows: set[str] = set()

    for d, date_id, date_label in date_id_pairs:
        params = {
            "SelectedSaleCategory": category_id,
            "SelectedSaleDate": date_id,
            "SelectedSaleStatus": status_id,
            "_": str(int(time.time() * 1000)),
        }
        try:
            resp = session.get(results_endpoint, params=params, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
        except requests.RequestException as e:
            log.warning("[%s] failed to fetch %s: %s", parish, date_label, e)
            continue
        time.sleep(REQUEST_DELAY_SECONDS)

        if "no results matching your criteria" in resp.text.lower():
            continue

        result_soup = BeautifulSoup(resp.text, "lxml")
        table = result_soup.find("table")
        if table is None:
            continue

        all_rows = table.find_all("tr")
        if not all_rows:
            continue

        # St. Tammany's table has no separate "Address" column.
        # The address is embedded inside the "Description" cell as text
        # below a property photo (confirmed via screenshot of live results).
        # Strategy: find the Description column by header, then extract the
        # street address from that cell's text using an address pattern.
        # Columns confirmed: List No | Case Number | Case Title | Description
        #                  | Attorney | Writ Amount | Terms & Conditions
        #                  | Sale Status | Details
        header_cells = all_rows[0].find_all(["th", "td"])
        desc_col_idx = None
        case_title_col_idx = None
        writ_col_idx = None
        for idx, h in enumerate(header_cells):
            label = h.get_text(strip=True).lower()
            if "description" in label:
                desc_col_idx = idx
            elif "case title" in label or "case style" in label:
                case_title_col_idx = idx
            elif "writ" in label or "amount" in label:
                writ_col_idx = idx

        # regex to pull a street address out of mixed cell text
        # matches patterns like: 250 LENWOOD DR, SLIDELL, LA 70458-1223
        ADDRESS_RE = re.compile(
            r"\d+\s+[A-Z0-9][A-Z0-9 .#&'-]{3,60},\s*[A-Z][A-Z .'-]{2,30},\s*LA\s+\d{5}(?:-\d{4})?",
            re.IGNORECASE
        )

        for row in all_rows[1:]:
            cells = row.find_all("td")
            if len(cells) < 2:
                continue
            row_text = row.get_text(" ", strip=True)
            if row_text in seen_rows:
                continue
            seen_rows.add(row_text)

            # price from Writ Amount column, or fallback scan of whole row
            if writ_col_idx is not None and writ_col_idx < len(cells):
                price = parse_price(cells[writ_col_idx].get_text(" ", strip=True))
            else:
                price = parse_price(row_text)

            # address: regex-extract from Description cell (address sits below
            # the property photo thumbnail as plain text)
            address = ""
            desc_cell_text = ""
            if desc_col_idx is not None and desc_col_idx < len(cells):
                desc_cell_text = cells[desc_col_idx].get_text(" ", strip=True)
                m = ADDRESS_RE.search(desc_cell_text)
                if m:
                    address = m.group(0).strip()

            # case title for description column
            if case_title_col_idx is not None and case_title_col_idx < len(cells):
                description = cells[case_title_col_idx].get_text(" ", strip=True)
            else:
                description = row_text

            all_listings.append(Listing(parish, address or "See description", price,
                                         date_label, description, base_url,
                                         property_type=classify_property_type(row_text),
                                         status="OK" if address else "NEEDS_FIX"))

    if not all_listings:
        return [Listing(parish, "", None, "", "Ran OK - 0 real estate sale(s) found across all upcoming dates",
                         base_url, status="NO_DATA")]
    return all_listings


def run_all_scrapers() -> list[Listing]:
    all_listings: list[Listing] = []

    for parish, county_id in VERIFIED_CIVILVIEW_PARISHES:
        log.info("Scraping %s (CivilView) ...", parish)
        try:
            all_listings.extend(scrape_civilview(parish, county_id))
        except Exception as e:
            log.error("Error scraping %s: %s", parish, e)
            all_listings.append(Listing(parish, "", None, "", str(e),
                                         f"https://salesweb.civilview.com/Sales/SalesSearch?countyId={county_id}",
                                         status="NEEDS_FIX"))

    log.info("Scraping Jefferson (JPSO) ...")
    try:
        all_listings.extend(scrape_jpso())
    except Exception as e:
        log.error("Error scraping Jefferson: %s", e)
        all_listings.append(Listing("Jefferson", "", None, "", str(e),
                                     "https://eservices2.jpso.com/JudpSale/Home/RealEstate",
                                     status="NEEDS_FIX"))

    all_listings.append(Listing("East Baton Rouge", "", None, "", EBR_NOTE,
                                 "https://foreclosure.ebrso.org/RealEstateSales/Index",
                                 status="NEEDS_JS"))

    log.info("Scraping St. Tammany ...")
    try:
        all_listings.extend(scrape_st_tammany())
    except Exception as e:
        log.error("Error scraping St. Tammany: %s", e)
        all_listings.append(Listing("St. Tammany", "", None, "", str(e),
                                     "https://public.stpso.com/Sheriff.PublicSite/",
                                     status="NEEDS_FIX"))

    for parish in UNVERIFIED_PARISHES:
        all_listings.append(Listing(parish, "", None, "", "Source not yet verified - see notes",
                                     "TBD", status="NOT_BUILT"))

    return all_listings


def write_excel(listings: list[Listing], out_path: str = "sheriff_sales_commercial.xlsx"):
    # Sort: Commercial first, then Land, Multi-Family, Residential, Unknown - cheapest first within each group
    type_order = {"Commercial": 0, "Land": 1, "Multi-Family": 2, "Residential": 3, "Unknown": 4}

    def sort_key(l: Listing):
        return (
            type_order.get(l.property_type, 4),
            0 if l.price is not None else 1,
            l.price if l.price is not None else float("inf"),
        )

    listings_sorted = sorted(listings, key=sort_key)

    # pre-fetch both free data sources once
    _load_redfin_zip_data()
    _load_census_zip_data()

    if RENTCAST_API_KEY:
        log.info("Rentcast API key detected - sqft lookups enabled (budget: %d calls this run).",
                  MAX_RENTCAST_CALLS_PER_RUN)
    else:
        log.info("No RENTCAST_API_KEY set - Square Footage / sqft-based Market Value will be "
                  "unavailable; Market Value will fall back to zip-level median comps.")

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheriff Sales"

    headers = [
        "Parish",                   # A
        "Property Type",            # B
        "Price",                    # C  ← writ/sale amount
        "Address",                  # D
        "Sale Date",                # E
        "Square Footage",           # F  ← Rentcast property-specific sqft (optional)
        "Avg $/sqft (Zip)",         # G  ← Redfin zip median (free)
        "Zip Median Sale Price",    # H  ← Redfin zip median sale price (free)
        "Census Median Value",      # I  ← Census ACS 5-yr estimate (free)
        "Market Value",             # J  ← sqft x psf when available, else zip-level fallback
        "Profit",                   # K  ← J − C
        "Description",              # L
        "Source URL",               # M
        "Status",                   # N
    ]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F5496")

    commercial_fill  = PatternFill("solid", fgColor="D9EAD3")
    land_fill        = PatternFill("solid", fgColor="FFF2CC")
    multifamily_fill = PatternFill("solid", fgColor="CFE2F3")
    residential_fill = PatternFill("solid", fgColor="F4CCCC")
    profit_fill      = PatternFill("solid", fgColor="E8F5E9")
    loss_fill        = PatternFill("solid", fgColor="FFEBEE")
    link_font        = Font(color="0563C1", underline="single")
    currency_fmt     = '"$"#,##0.00'
    integer_fmt      = '#,##0'

    total = len([l for l in listings_sorted if l.status == "OK"])
    log.info("Enriching %d listings with market data ...", total)

    ADDRESS_COL = 4
    SOURCE_COL = 13

    for l in listings_sorted:
        row_idx = ws.max_row + 1
        has_real_address = bool(l.address) and l.address not in (
            "Address not available", "See case details", "Unknown address", "See description"
        )

        # --- market data enrichment ---
        sqft = avg_psf = median_sale = census_val = market_value = profit = None
        if has_real_address and l.status == "OK":
            mkt = get_zip_market_data(l.address)
            avg_psf      = mkt.get("psf")
            median_sale  = mkt.get("median_sale_price")
            census_val   = mkt.get("census_value")

            sqft = get_property_sqft(l.address)
            if sqft and avg_psf:
                # property-specific estimate, as documented at the top of this file
                market_value = round(sqft * avg_psf, 2)
            else:
                # fall back to the zip-level comp value when we don't have a
                # sqft figure (no Rentcast key, lookup failed, budget hit, etc.)
                market_value = mkt.get("market_value")

            if market_value is not None and l.price is not None:
                profit = round(market_value - l.price, 2)

        ws.append([
            l.parish, l.property_type, l.price,
            l.address, l.sale_date,
            sqft, avg_psf, median_sale, census_val, market_value, profit,
            l.description, l.source_url, l.status,
        ])

        # --- formatting ---
        # address hyperlink
        address_cell = ws.cell(row=row_idx, column=ADDRESS_COL)
        if has_real_address:
            address_cell.hyperlink = "https://www.google.com/maps/search/" + quote_plus(l.address)
            address_cell.font = link_font

        # source URL hyperlink
        source_cell = ws.cell(row=row_idx, column=SOURCE_COL)
        if l.source_url and l.source_url != "TBD":
            source_cell.hyperlink = l.source_url
            source_cell.font = link_font

        # number formats
        # C=3 Price, F=6 Sqft, G=7 Avg$/sqft, H=8 Median Sale, I=9 Census, J=10 Market Value, K=11 Profit
        ws.cell(row=row_idx, column=6).number_format = integer_fmt
        for col in (3, 7, 8, 9, 10, 11):
            ws.cell(row=row_idx, column=col).number_format = currency_fmt

        # row background by property type
        fill_map = {
            "Commercial":   commercial_fill,
            "Land":         land_fill,
            "Multi-Family": multifamily_fill,
            "Residential":  residential_fill,
        }
        row_fill = fill_map.get(l.property_type)
        if row_fill:
            for col in range(1, len(headers) + 1):
                if col not in (ADDRESS_COL, SOURCE_COL):  # don't override hyperlink font cells
                    ws.cell(row=row_idx, column=col).fill = row_fill
            address_cell.fill = row_fill
            source_cell.fill = row_fill

        # highlight profit cell green/red
        profit_cell = ws.cell(row=row_idx, column=11)
        if profit is not None:
            profit_cell.fill = profit_fill if profit >= 0 else loss_fill
            profit_cell.font = Font(bold=True,
                                     color="1B5E20" if profit >= 0 else "B71C1C")

    # column widths: A  B  C  D  E  F  G  H  I  J  K  L  M  N
    col_widths =    [14,14,14,32,12,11,14,18,18,16,16,50,40,12]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    wb.save(out_path)
    log.info("Saved %d listings to %s", len(listings_sorted), out_path)


if __name__ == "__main__":
    results = run_all_scrapers()
    write_excel(results)
    needs_fix = sum(1 for l in results if l.status != "OK")
    log.info("Done. %d/%d rows need selector fixes - check 'Status' column.", needs_fix, len(results))
