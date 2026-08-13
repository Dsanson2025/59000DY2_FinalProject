# Louisiana Sheriff Sale Scraper

Scrapes upcoming sheriff sale / foreclosure listings for several Louisiana
metro parishes and outputs a single formatted Excel workbook — classified
by property type, enriched with free zip-level market comps, and (optionally)
an estimated Market Value and Profit for each listing.

## What it does

For every upcoming sale it can reach, the script pulls the case number, sale
date, address, writ/sale price, and case description, then:

1. **Classifies the property type** — `Commercial`, `Land`, `Multi-Family`,
   `Residential`, or `Unknown` — using keyword matching plus a heuristic that
   flags a sale as Commercial when the defendant name looks like a business
   entity (LLC, Inc., Corp., etc.) rather than a person.
2. **Looks up free zip-level market data** — median $/sqft, median sale
   price, and median list price from Redfin's public market tracker, plus
   Census Bureau ACS median home values, keyed off the zip code in the
   property address.
3. **Optionally looks up the property's actual square footage** via the
   Rentcast API (free tier, see [Optional: square footage lookups](#optional-square-footage-lookups-rentcast)
   below) and uses `sqft × zip $/sqft` as a more precise Market Value
   estimate. Without a Rentcast key, Market Value falls back to the zip's
   median comp value.
4. **Estimates Profit** as `Market Value − Writ/Sale Price`.
5. Writes everything to `sheriff_sales_commercial.xlsx`, sorted with
   Commercial listings first (then Land, Multi-Family, Residential, Unknown),
   cheapest first within each group, with color-coded rows by property type,
   clickable address (Google Maps) and source links, and green/red
   highlighting on the Profit column.

## Parish coverage

| Parish | Status | Notes |
|---|---|---|
| Orleans | ✅ Verified | CivilView SalesWeb (`countyId=28`) |
| Ascension | ✅ Verified | CivilView SalesWeb (`countyId=55`) |
| Jefferson | ✅ Verified | JPSO real estate sales list (auto-discovers the date-picker form) |
| St. Tammany | ✅ Verified | Reverse-engineered AJAX endpoint (`RealEstateMovablesSearchResultsTable`) |
| East Baton Rouge | ⚠️ Needs work | Results table is JS/AJAX-rendered; the underlying JSON endpoint hasn't been identified yet. See the `EBR_NOTE` comment in the script for how to find it via Chrome DevTools. |
| Caddo, Calcasieu, Lafayette, Ouachita, Rapides, Terrebonne, Tangipahoa, Bossier | ❌ Not built | No working scrape source confirmed yet — several require phone calls, PDF legal notices, or paid platforms (e.g. Bid4Assets) that block scraping. These still appear in the output as `NOT_BUILT` rows so the gap is visible rather than silent. |

A row's **Status** column tells you what happened: `OK` (parsed fine),
`NEEDS_FIX` (page structure changed or a field couldn't be found),
`NO_DATA` (ran fine, just nothing upcoming), or `NOT_BUILT` (no scraper
exists yet for that parish).

## Requirements

- Python 3.9+
- Install dependencies:

  ```bash
  pip install requests beautifulsoup4 openpyxl lxml
  ```

## Optional: square footage lookups (Rentcast)

Without any setup, the script works out of the box using only free data
(Redfin + Census) and estimates Market Value from zip-level comps. To get a
more precise, property-specific Market Value based on actual square footage:

1. Sign up free at [app.rentcast.io](https://app.rentcast.io) (50 requests/month on the free tier).
2. Grab your API key from [app.rentcast.io/app/api-keys](https://app.rentcast.io/app/api-keys).
3. Set it as an environment variable before running:

   ```bash
   # Mac/Linux
   export RENTCAST_API_KEY=your_key_here

   # Windows
   set RENTCAST_API_KEY=your_key_here
   ```

   (Or paste it directly into the `RENTCAST_API_KEY` line near the top of
   the script.)

The script caps itself at 45 Rentcast calls per run so a single run can't
exhaust the whole month's free-tier budget by itself; once the cap is hit
(or for any address Rentcast can't match), it just falls back to the
zip-level comp estimate instead of failing.

## Usage

```bash
python sheriff_sale_scraper.py
```

This scrapes every verified parish's upcoming sale dates, enriches each
listing with market data, and writes `sheriff_sales_commercial.xlsx` to the
current folder. A run typically takes a few minutes — it deliberately waits
~1.5 seconds between requests to each site to avoid hammering small
government servers, and the first run also streams Redfin's ~1.5GB public
market-tracker file (filtered down to Louisiana zips on the fly, so it
doesn't load the whole thing into memory).

## Output columns

| Column | Description |
|---|---|
| Parish | Which parish the sale belongs to |
| Property Type | Commercial / Land / Multi-Family / Residential / Unknown |
| Price | Writ or sale amount |
| Address | Property address (click-through to Google Maps) |
| Sale Date | Scheduled sheriff sale date |
| Square Footage | Property-specific sqft from Rentcast (blank if no API key or no match) |
| Avg $/sqft (Zip) | Redfin median $/sqft for that zip |
| Zip Median Sale Price | Redfin median sale price for that zip |
| Census Median Value | Census ACS 5-year median home value for that zip |
| Market Value | `sqft × Avg $/sqft` when available, else the zip-level fallback |
| Profit | `Market Value − Price` (green if positive, red if negative) |
| Description | Case title / listing description |
| Source URL | Link back to the source sheriff's office page |
| Status | `OK`, `NEEDS_FIX`, `NO_DATA`, or `NOT_BUILT` |

## Extending to a new parish

Most Louisiana sheriff/constable sale pages fall into one of a few patterns
already handled here:

- **CivilView SalesWeb** (`scrape_civilview`) — used by several parishes;
  just needs the parish's `countyId`.
- **Custom ASP.NET form with a date dropdown** (`scrape_jpso`) — auto-discovers
  the form fields, so it can often be adapted with just a new base URL.
- **AJAX/JSON endpoint** (`scrape_st_tammany`) — for sites whose results
  table is loaded client-side; find the underlying request in Chrome
  DevTools → Network tab, then mirror that pattern.
- **Anything else** — `scrape_generic_sheriff_site()` is a best-effort
  fallback: pass it a URL and a CSS selector for each sale row, and it will
  fall back further to scanning the raw page text for `$` amounts and
  property-type keywords if no selector is supplied.

## A note on responsible use

This script only reads publicly posted sheriff sale listings and adds
built-in delays between requests. Before pointing it at a new site, check
that site's terms of use/robots.txt, and keep the request delay in place —
these are small government servers, not commercial APIs.
