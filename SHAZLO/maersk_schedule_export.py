"""
maersk_schedule_export.py

Fetches Maersk "Point to Point" shipping schedules for a lane across a date
range and exports every sailing found to an Excel file.

Pre-configured for: Shanghai, China (CNSGH)  ->  Manzanillo, Mexico (MXZLO)
Date range: September 2026 - November 2026

Requirements (install once):
    pip install playwright beautifulsoup4 pandas openpyxl
    playwright install chromium

Run:
    python maersk_schedule_export.py

Notes
-----
* This script must be run from a machine with normal internet access to
  www.maersk.com (it will NOT work from a network-restricted sandbox).
* It drives a real (headless) Chromium browser via Playwright rather than
  plain `requests`, because maersk.com's schedule page is a JS application
  and simple HTTP GET requests are commonly blocked by its bot protection.
* The page/URL parameters below (from, to, fromRkstCode, toRkstCode, ...)
  were captured from the live "Shanghai -> Manzanillo" schedule search on
  maersk.com. If Maersk changes its site, you may need to redo the search
  in a browser once and copy the new URL parameters in below.
* The Maersk site paginates results in ~4-week windows ("numberOfWeeks=4"),
  so this script requests one window at a time and stitches the results
  together, de-duplicating by (departure date, vessel, voyage).
* Vessel name and voyage number are rendered back-to-back with no space in
  between (e.g. "MAERSK NARVIK636S", "SHIMANAMI BAYO4ES"), and splitting
  them reliably isn't possible from the text alone (voyage code lengths and
  formats vary by service). Rather than risk a silently wrong split, this
  script keeps them together in one "Vessel / Voyage" column exactly as
  shown on the site.
"""

from datetime import date, timedelta

import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# --------------------------------------------------------------------------
# Configuration - change these to search a different lane / date range
# --------------------------------------------------------------------------

# Location codes as used by maersk.com for this lane (Shanghai -> Manzanillo)
FROM_LOCODE = "2IW9P6J7XAW72"   # Shanghai, China
TO_LOCODE = "3AJMN274REBUN"     # Manzanillo, Mexico
FROM_RKST_CODE = "CNSGH"
TO_RKST_CODE = "MXZLO"

CONTAINER_ISO_CODE = "45G1"     # 40ft High Cube dry container (site default)
FROM_SERVICE_MODE = "CY"        # Container Yard (port-to-port)
TO_SERVICE_MODE = "CY"

# Search window
SEARCH_START = date(2026, 9, 1)
SEARCH_END = date(2026, 11, 30)

OUTPUT_FILE = "maersk_shanghai_to_manzanillo_sep_nov_2026.xlsx"

BASE_URL = "https://www.maersk.com/schedules/pointToPoint"
WINDOW_WEEKS = 4  # matches the window size the site itself uses

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
)

# --------------------------------------------------------------------------
# URL building
# --------------------------------------------------------------------------


def build_url(window_start: date, weeks: int = WINDOW_WEEKS) -> str:
    window_end = window_start + timedelta(weeks=weeks)
    params = {
        "from": FROM_LOCODE,
        "to": TO_LOCODE,
        "fromRkstCode": FROM_RKST_CODE,
        "toRkstCode": TO_RKST_CODE,
        "containerIsoCode": CONTAINER_ISO_CODE,
        "fromServiceMode": FROM_SERVICE_MODE,
        "toServiceMode": TO_SERVICE_MODE,
        "numberOfWeeks": str(weeks),
        "dateType": "D",
        "date": window_start.isoformat(),
        "dateTo": window_end.isoformat(),
        "vesselFlag": "",
    }
    query = "&".join(f"{k}={v}" for k, v in params.items())
    return f"{BASE_URL}?{query}"


def date_windows(start: date, end: date, weeks: int = WINDOW_WEEKS):
    """Yield window start dates stepping by `weeks`, covering [start, end]."""
    cur = start
    while cur <= end:
        yield cur
        cur += timedelta(weeks=weeks)


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------

DEADLINE_LABELS = {
    "Empty container pickup",
    "Container gate-in",
    "Shipping Instructions",
    "Verified Gross Mass",
    "Dangerous Goods Declaration",
}


def parse_schedule(html: str):
    """Parse sailing cards out of a rendered schedule page's visible text.

    Maersk doesn't expose a documented JSON API for this page, and its CSS
    class names are auto-generated/unstable, so this walks the page's
    visible text line by line looking for the repeating "Departure ...
    Arrival ... Vessel/Voyage ... Transit Time:" pattern that each sailing
    card renders as. The vessel/voyage section is read flexibly (1 or more
    lines, until the "Transit Time:" line is hit) since depending on the
    page markup, the vessel name and voyage code can end up as one merged
    line or two separate lines.
    """
    soup = BeautifulSoup(html, "html.parser")
    main = soup.find("main") or soup
    text = main.get_text("\n", strip=True)
    lines = [line for line in text.split("\n") if line]

    results = []
    i = 0
    n = len(lines)
    while i < n:
        if (
            lines[i] == "Departure"
            and i + 6 < n
            and lines[i + 3] == "Arrival"
            and lines[i + 6] == "Vessel/Voyage"
        ):
            dep_date, dep_port = lines[i + 1], lines[i + 2]
            arr_date, arr_port = lines[i + 4], lines[i + 5]

            # Gather the vessel/voyage line(s) up to "Transit Time:".
            j = i + 7
            vv_lines = []
            while j < n and not lines[j].startswith("Transit Time:"):
                vv_lines.append(lines[j])
                j += 1
                if len(vv_lines) > 3:  # safety valve against malformed pages
                    break
            if j >= n or not lines[j].startswith("Transit Time:"):
                i += 1
                continue
            transit = lines[j].replace("Transit Time:", "").strip()

            vessel_voyage = " ".join(vv_lines) if vv_lines else ""

            row = {
                "Departure Date": dep_date,
                "Departure Port": dep_port,
                "Arrival Date": arr_date,
                "Arrival Port": arr_port,
                "Vessel / Voyage": vessel_voyage,
                "Transit Time": transit,
                "Empty container pickup": "",
                "Container gate-in": "",
                "Shipping Instructions": "",
                "Verified Gross Mass": "",
                "Dangerous Goods Declaration": "",
            }

            # Optionally pick up the "Deadlines" block that follows, up to
            # the next "Departure" card or "Show route details".
            k = j + 1
            while k < n and lines[k] != "Departure" and lines[k] != "Show route details":
                if lines[k] in DEADLINE_LABELS and k + 1 < n:
                    row[lines[k]] = lines[k + 1]
                    k += 2
                    continue
                k += 1

            results.append(row)
            i = k
            continue
        i += 1
    return results


# --------------------------------------------------------------------------
# Fetching
# --------------------------------------------------------------------------


def fetch_range(start: date, end: date):
    all_rows = []
    seen = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=USER_AGENT)

        for window_start in date_windows(start, end):
            url = build_url(window_start)
            print(f"Fetching window starting {window_start} ...")
            page.goto(url, wait_until="networkidle", timeout=60000)
            page.wait_for_timeout(2000)  # let any client-side render settle
            html = page.content()

            rows = parse_schedule(html)
            for row in rows:
                key = (row["Departure Date"], row["Vessel / Voyage"])
                if key not in seen:
                    seen.add(key)
                    all_rows.append(row)

        browser.close()

    return all_rows


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------


def main():
    rows = fetch_range(SEARCH_START, SEARCH_END)

    if not rows:
        print(
            "No sailings were parsed. Maersk may have changed its page "
            "layout, or bot protection blocked the request - try running "
            "with headless=False in fetch_range() to see what loaded."
        )
        return

    df = pd.DataFrame(rows)
    df["_dep_parsed"] = pd.to_datetime(
        df["Departure Date"], format="%d %b %Y", errors="coerce"
    )
    df = df[
        (df["_dep_parsed"] >= pd.Timestamp(SEARCH_START))
        & (df["_dep_parsed"] <= pd.Timestamp(SEARCH_END))
    ]
    df = df.sort_values("_dep_parsed").drop(columns=["_dep_parsed"])

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(df)} sailings to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
