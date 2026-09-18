#!/usr/bin/env python3
"""
hmm_schedule_scraper.py
========================

Scrapes the HMM (Hyundai Merchant Marine) e-Service "Point to Point" sailing
schedule page for a given origin/destination pair and date range, and saves
the results to an Excel (.xlsx) and CSV file.

Website: https://www.hmm21.com/e-service/general/schedule/ScheduleMainPost.do

Default lane / period (edit the CONFIG section below to change):
    Origin      : SHANGHAI, CHINA   (CNSHA)
    Destination : HOCHIMINH, VIETNAM (VNSGN)
    Period      : 2026-09-01 -> 2026-11-30

WHY BROWSER AUTOMATION (Playwright) INSTEAD OF PLAIN HTTP REQUESTS
--------------------------------------------------------------------
HMM's site routes its AJAX search calls through a bot-protection layer that
rewrites every XHR endpoint to an obfuscated, session-specific URL (it is
not a stable REST API you can call directly with `requests`). The reliable
way to automate this page is to drive a real browser, exactly like a human
would: fill in the form, click "Retrieve", and read the rendered results.

HOW IT WORKS
------------
1. Opens the schedule page in a Chromium browser (Playwright).
2. Fills the Origin / Destination autocomplete fields.
3. The "Sailing date (by Vessel)" search only returns a window of 2-8 weeks
   at a time, so to cover a multi-month period the script runs the search
   multiple times with a sliding start date (with a little overlap) and
   removes duplicate sailings by voyage code + ETD.
4. Parses each result card's text (week, ETD/ETA, ports, transit time,
   routing, vessel, route/loop, operator, cut-off times).
5. Filters to sailings whose ETD falls inside the requested period and
   writes them out sorted by ETD.

REQUIREMENTS
------------
    pip install playwright pandas openpyxl --break-system-packages
    playwright install chromium      # first time only, downloads the browser

USAGE
-----
    python hmm_schedule_scraper.py
    python hmm_schedule_scraper.py --origin SHANGHAI --destination HOCHIMINH \
        --start 2026-09-01 --end 2026-11-30 --headed

Run with --headed (or edit HEADLESS below) if the site's bot-protection
blocks the default headless mode - some sites only challenge headless
browsers.
"""

import argparse
import json
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except ImportError:
    sys.exit(
        "Playwright is not installed.\n"
        "Install it with:\n"
        "    pip install playwright pandas openpyxl --break-system-packages\n"
        "    playwright install chromium\n"
    )

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas is required. Install with: pip install pandas openpyxl --break-system-packages")


# ----------------------------------------------------------------------------
# CONFIG - edit these defaults if you don't want to pass command-line flags
# ----------------------------------------------------------------------------
SCHEDULE_URL = "https://www.hmm21.com/e-service/general/schedule/ScheduleMainPost.do"

DEFAULT_ORIGIN = "SHANGHAI"          # text typed into the Origin autocomplete box
DEFAULT_ORIGIN_CODE = "CNSHA"        # UN/LOCODE shown in the autocomplete suggestion
DEFAULT_DESTINATION = "HOCHIMINH"    # text typed into the Destination autocomplete box
DEFAULT_DESTINATION_CODE = "VNSGN"   # UN/LOCODE shown in the autocomplete suggestion

DEFAULT_START = "2026-09-01"
DEFAULT_END = "2026-11-30"

HEADLESS = True          # set False (or use --headed) if the site blocks headless browsers
MAX_WEEKS_PER_QUERY = 8   # the site's dropdown tops out at 8 weeks per search
OVERLAP_DAYS = 5          # small overlap between consecutive queries so nothing is missed

OUTPUT_DIR = Path(__file__).resolve().parent


# ----------------------------------------------------------------------------
# Parsing
# ----------------------------------------------------------------------------
# Each result card's compact text looks like:
#   "WK 45 Origin 2026-11-07 SHANGHAI,CHINA 4 days Direct Destination
#    2026-11-12 HOCHIMINH, VIETNAM Main Vessel POS LAEMCHABANG (1046W)
#    Route VTX Operator PAN 1st Vessel POS LAEMCHABANG (1046W)
#    Inland Cut-off 2026-11-05 21:00 Port Cut-off 2026-11-05 21:00
#    DOC Cut-off 2026-11-05 21:00 Book via Book Now Show Details"
RESULT_PATTERN = re.compile(
    r"WK\s*(?P<week>\d+)\s*"
    r"Origin\s+(?P<etd>\d{4}-\d{2}-\d{2})\s+(?P<origin_port>.*?)\s+"
    r"(?P<transit_days>\d+)\s*days?\s+(?P<routing>Direct|T/S)\s*"
    r"Destination\s+(?P<eta>\d{4}-\d{2}-\d{2})\s+(?P<dest_port>.*?)\s+"
    r"Main Vessel\s+(?P<main_vessel>.*?)\s+Route\s+(?P<route>\S+)\s+"
    r"Operator\s+(?P<operator>\S+)\s+"
    r"1st Vessel\s+(?P<first_vessel>.*?)\s+"
    r"Inland Cut-off\s+(?P<inland_cutoff>\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})\s+"
    r"Port Cut-off\s+(?P<port_cutoff>\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})\s+"
    r"DOC Cut-off\s+(?P<doc_cutoff>\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})"
)


def parse_result_text(text: str) -> dict | None:
    """Parse one result card's flattened text into a dict of fields."""
    text = re.sub(r"\s+", " ", text).strip()
    m = RESULT_PATTERN.search(text)
    if not m:
        return None
    d = m.groupdict()
    d["week"] = int(d["week"])
    d["transit_days"] = int(d["transit_days"])
    return d


# ----------------------------------------------------------------------------
# Date-window planning
# ----------------------------------------------------------------------------
def build_windows(start: date, end: date, max_weeks: int, overlap_days: int):
    """Split [start, end] into overlapping (window_start, weeks) chunks that
    respect the site's 2-8 week search limit."""
    windows = []
    cursor = start
    max_days = max_weeks * 7
    while cursor <= end:
        windows.append((cursor, max_weeks))
        step = max_days - overlap_days
        cursor = cursor + timedelta(days=step)
    return windows


# ----------------------------------------------------------------------------
# Browser automation
# ----------------------------------------------------------------------------
def dismiss_popups(page):
    """Best-effort dismissal of cookie/consent banners or chat widgets that
    could intercept clicks."""
    for selector in [
        "text=Accept",
        "text=I Agree",
        "button:has-text('Close')",
        ".cookie-consent button",
    ]:
        try:
            el = page.locator(selector).first
            if el.is_visible(timeout=1000):
                el.click(timeout=1000)
        except Exception:
            pass


def select_autocomplete(page, input_id: str, query: str, expect_code: str | None = None):
    """Type into an autocomplete field (#srchPointFrom / #srchPointTo) and
    click the matching suggestion."""
    field = page.locator(f"#{input_id}")
    field.click()
    field.fill("")
    field.type(query, delay=50)
    page.wait_for_selector("div.ac_results li", timeout=10000)
    page.wait_for_timeout(300)  # let the suggestion list settle

    suggestions = page.locator("div.ac_results li")
    count = suggestions.count()
    chosen = None
    for i in range(count):
        item = suggestions.nth(i)
        item_text = item.inner_text()
        if expect_code and expect_code in item_text:
            chosen = item
            break
        if chosen is None and query.split()[0].upper() in item_text.upper():
            chosen = item
    if chosen is None:
        chosen = suggestions.first
    chosen.click()
    page.wait_for_timeout(300)


def run_search(page, window_start: date, weeks: int):
    """Set the sailing-date window and click Retrieve; wait for the result
    list to refresh."""
    try:
        old_count = page.locator("#resultCount").inner_text()
    except Exception:
        old_count = None

    page.fill("#srchSailDate", window_start.strftime("%Y-%m-%d"))
    page.select_option("#srchSelWeeks", value=str(weeks))
    page.click("#btnRetrieve")

    if old_count is not None:
        try:
            page.wait_for_function(
                """(old) => {
                    const el = document.querySelector('#resultCount');
                    return el && el.innerText !== old;
                }""",
                arg=old_count,
                timeout=20000,
            )
        except PlaywrightTimeoutError:
            pass
    page.wait_for_timeout(1200)


def scrape_current_results(page, window_start: date, weeks: int) -> list[dict]:
    texts = page.eval_on_selector_all(
        "#lsitContentArea2 .info-list-area > ul > li .result-info .list-result",
        "elements => elements.map(e => e.innerText)",
    )
    records = []
    for t in texts:
        parsed = parse_result_text(t)
        if parsed:
            parsed["_query_window_start"] = window_start.isoformat()
            parsed["_query_weeks"] = weeks
            records.append(parsed)
        else:
            print(f"  ! could not parse a result card: {t[:120]!r}", file=sys.stderr)
    return records


def scrape_schedule(
    origin: str,
    origin_code: str,
    destination: str,
    destination_code: str,
    start: date,
    end: date,
    headless: bool = HEADLESS,
) -> list[dict]:
    windows = build_windows(start, end, MAX_WEEKS_PER_QUERY, OVERLAP_DAYS)
    all_records: list[dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()
        page.goto(SCHEDULE_URL, wait_until="load", timeout=60000)
        dismiss_popups(page)

        print(f"Setting Origin -> {origin} ({origin_code})")
        select_autocomplete(page, "srchPointFrom", origin, origin_code)
        print(f"Setting Destination -> {destination} ({destination_code})")
        select_autocomplete(page, "srchPointTo", destination, destination_code)

        for window_start, weeks in windows:
            print(f"Retrieving {weeks}-week window starting {window_start} ...")
            run_search(page, window_start, weeks)
            records = scrape_current_results(page, window_start, weeks)
            print(f"  -> {len(records)} sailings parsed")
            all_records.extend(records)

        browser.close()

    return all_records


# ----------------------------------------------------------------------------
# Post-processing
# ----------------------------------------------------------------------------
def dedupe_and_filter(records: list[dict], start: date, end: date) -> "pd.DataFrame":
    df = pd.DataFrame(records)
    if df.empty:
        return df

    df["etd_date"] = pd.to_datetime(df["etd"]).dt.date
    df["eta_date"] = pd.to_datetime(df["eta"]).dt.date

    # Same sailing can appear in more than one overlapping query window.
    df = df.drop_duplicates(subset=["main_vessel", "route", "etd", "eta"])

    df = df[(df["etd_date"] >= start) & (df["etd_date"] <= end)]
    df = df.sort_values("etd_date").reset_index(drop=True)

    columns = [
        "week", "etd", "origin_port", "eta", "dest_port", "transit_days",
        "routing", "main_vessel", "first_vessel", "route", "operator",
        "inland_cutoff", "port_cutoff", "doc_cutoff",
    ]
    return df[[c for c in columns if c in df.columns]]


def save_outputs(df: "pd.DataFrame", origin: str, destination: str, start: date, end: date):
    stem = f"HMM_schedule_{origin}_{destination}_{start}_{end}".replace(" ", "")
    xlsx_path = OUTPUT_DIR / f"{stem}.xlsx"
    csv_path = OUTPUT_DIR / f"{stem}.csv"

    df.to_excel(xlsx_path, index=False, sheet_name="Schedule")
    df.to_csv(csv_path, index=False)

    print(f"\nSaved {len(df)} sailings to:")
    print(f"  {xlsx_path}")
    print(f"  {csv_path}")


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------
def parse_args():
    parser = argparse.ArgumentParser(description="Scrape HMM sailing schedule (Point to Point).")
    parser.add_argument("--origin", default=DEFAULT_ORIGIN, help="Origin search text, e.g. SHANGHAI")
    parser.add_argument("--origin-code", default=DEFAULT_ORIGIN_CODE, help="Origin UN/LOCODE, e.g. CNSHA")
    parser.add_argument("--destination", default=DEFAULT_DESTINATION, help="Destination search text, e.g. HOCHIMINH")
    parser.add_argument("--destination-code", default=DEFAULT_DESTINATION_CODE, help="Destination UN/LOCODE, e.g. VNSGN")
    parser.add_argument("--start", default=DEFAULT_START, help="Start date YYYY-MM-DD")
    parser.add_argument("--end", default=DEFAULT_END, help="End date YYYY-MM-DD")
    parser.add_argument("--headed", action="store_true", help="Run with a visible browser window")
    return parser.parse_args()


def main():
    args = parse_args()
    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    end = datetime.strptime(args.end, "%Y-%m-%d").date()

    records = scrape_schedule(
        origin=args.origin,
        origin_code=args.origin_code,
        destination=args.destination,
        destination_code=args.destination_code,
        start=start,
        end=end,
        headless=not args.headed,
    )

    df = dedupe_and_filter(records, start, end)
    if df.empty:
        print("No sailings found / parsed. The page layout may have changed, "
              "or the site blocked automated access - try running with --headed.")
        return

    save_outputs(df, args.origin, args.destination, start, end)


if __name__ == "__main__":
    main()
