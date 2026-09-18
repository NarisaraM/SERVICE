#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
COSCO SHIPPING Lines - Sailing Schedule fetcher
Route: Shanghai (CNSHA) -> Ho Chi Minh / Cat Lai (VNCAL)
Period: default September - November (current year, editable below)

How it works
------------
Instead of driving a real browser, this script calls the same JSON API that
the COSCO eLines website (elines.coscoshipping.com) calls internally when you
fill in the "Find Schedule by City Pairs" form and click "Search":

    POST https://elines.coscoshipping.com/ebschedule/public/purpoShipmentWs

The endpoint is public (no login required) and returns the full list of
scheduled sailings for the given origin/destination/date range as JSON,
which is exactly what is shown in the schedule table on the website.

Requirements
------------
    pip install requests pandas openpyxl

Usage
-----
    python cosco_schedule_shanghai_hochiminh.py
    python cosco_schedule_shanghai_hochiminh.py --from 2026-09-01 --to 2026-11-30
"""

import argparse
import datetime as dt
import sys

import requests

try:
    import pandas as pd
except ImportError:
    pd = None


API_URL = "https://elines.coscoshipping.com/ebschedule/public/purpoShipmentWs"

# City identifiers as used internally by the COSCO eLines website for
# Shanghai and Ho Chi Minh (Cat Lai terminal). These were captured from the
# live site and are stable (they are the site's own database IDs for the
# city pair).
ORIGIN_CITY_UUID = "738872886232873"
ORIGIN_CITY = "Shanghai,Shanghai,Shanghai,China,CNSHA"

DESTINATION_CITY_UUID = "882736036768074"
DESTINATION_CITY = "Ho Chi Minh (Cat Lai), ,Ho Chi Minh,Vietnam,VNCAL"

HEADERS = {
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
    ),
    "Origin": "https://elines.coscoshipping.com",
    "Referer": "https://elines.coscoshipping.com/ebusiness/sailingSchedule/searchByCity/resultByCity",
}


def fetch_schedule(from_date: str, to_date: str, cargo_nature: str = "") -> list[dict]:
    """Call the COSCO schedule API and return the raw list of voyage legs.

    from_date / to_date: strings in "YYYY-MM-DD" format.
    """
    payload = {
        "fromDate": from_date,
        "pickup": "B",       # B = BOTH (CY/DOOR) for OB haulage, matches "BOTH" dropdown
        "delivery": "B",     # B = BOTH for IB haulage
        "estimateDate": "D", # D = filter by (earliest) Departure Date
        "toDate": to_date,
        "originCityUuid": ORIGIN_CITY_UUID,
        "destinationCityUuid": DESTINATION_CITY_UUID,
        "originCity": ORIGIN_CITY,
        "destinationCity": DESTINATION_CITY,
        "cargoNature": cargo_nature,
    }

    resp = requests.post(API_URL, json=payload, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    body = resp.json()

    if str(body.get("code")) != "200":
        raise RuntimeError(f"COSCO API returned an error: {body}")

    rows = body.get("data", {}).get("content", {}).get("data", []) or []
    return rows


def to_dataframe(rows: list[dict]):
    """Turn the raw API rows into a clean, sorted table of the useful columns."""
    if pd is None:
        raise RuntimeError("pandas is required for to_dataframe(); pip install pandas openpyxl")

    if not rows:
        return pd.DataFrame(
            columns=[
                "Vessel", "Voyage", "Service", "Cargo Nature", "Cut Off",
                "POL", "ETD", "POD", "ETA", "Transit (days)", "Available",
                "OB Haulage", "IB Haulage",
            ]
        )

    df = pd.DataFrame(rows)

    rename_map = {
        "vessel": "Vessel",
        "extVoyage": "Voyage",
        "service": "Service",
        "cargoNature": "Cargo Nature",
        "cutOff": "Cut Off",
        "pol": "POL",
        "etd": "ETD",
        "pod": "POD",
        "eta": "ETA",
        "transitTime": "Transit (days)",
        "available": "Available",
        "outboundHaulage": "OB Haulage",
        "inboundHaulage": "IB Haulage",
    }
    cols = [c for c in rename_map if c in df.columns]
    out = df[cols].rename(columns=rename_map)

    if "ETD" in out.columns:
        out["_etd_sort"] = pd.to_datetime(out["ETD"], errors="coerce")
        out = out.sort_values("_etd_sort").drop(columns="_etd_sort")

    return out.reset_index(drop=True)


def default_date_range() -> tuple[str, str]:
    """September 1 to November 30 of the current year."""
    year = dt.date.today().year
    return f"{year}-09-01", f"{year}-11-30"


def main():
    default_from, default_to = default_date_range()

    parser = argparse.ArgumentParser(
        description="Fetch COSCO SHIPPING Lines sailing schedule: Shanghai -> Ho Chi Minh"
    )
    parser.add_argument("--from", dest="from_date", default=default_from,
                         help=f"Earliest departure date, YYYY-MM-DD (default {default_from})")
    parser.add_argument("--to", dest="to_date", default=default_to,
                         help=f"Latest departure date, YYYY-MM-DD (default {default_to})")
    parser.add_argument("--cargo-nature", dest="cargo_nature", default="",
                         help='Optional cargo nature filter, e.g. "GC" or "RF" (default: all)')
    parser.add_argument("--out", dest="out_file", default="cosco_schedule.xlsx",
                         help="Output Excel/CSV filename (default cosco_schedule.xlsx)")
    args = parser.parse_args()

    print(f"Fetching COSCO schedule: Shanghai -> Ho Chi Minh, "
          f"{args.from_date} to {args.to_date} ...")

    try:
        rows = fetch_schedule(args.from_date, args.to_date, args.cargo_nature)
    except requests.RequestException as e:
        print(f"Network/request error: {e}", file=sys.stderr)
        sys.exit(1)
    except RuntimeError as e:
        print(str(e), file=sys.stderr)
        sys.exit(1)

    print(f"Found {len(rows)} voyage leg(s).")

    if pd is not None:
        df = to_dataframe(rows)
        with pd.option_context("display.max_rows", None, "display.width", 160):
            print(df.to_string(index=False))

        if args.out_file.lower().endswith(".csv"):
            df.to_csv(args.out_file, index=False, encoding="utf-8-sig")
        else:
            df.to_excel(args.out_file, index=False)
        print(f"\nSaved: {args.out_file}")
    else:
        # Fallback without pandas: just dump the raw rows.
        import json
        print(json.dumps(rows, indent=2, ensure_ascii=False))
        print("\n(pandas not installed - install it to get a table + Excel export: "
              "pip install pandas openpyxl)")


if __name__ == "__main__":
    main()
