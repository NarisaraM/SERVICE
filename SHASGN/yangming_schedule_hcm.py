"""
Yang Ming Point-to-Point Schedule Fetcher
==========================================
Shanghai (CNSHA) -> Ho Chi Minh City (VNSGN)

Yang Ming's "Point-to-Point Search" page (yangming.com/en/esolution/schedule/
point_to_point_search) calls a public JSON API under the hood:

    GET https://www.yangming.com/api/P2P/GetP2PRoutes
        ?locationCodeFrom=CNSHA&serviceTermFrom=Y
        &locationCodeTo=VNSGN&serviceTermTo=Y
        &priorityWay=ALL&dateDefinition=DEP
        &startDate=YYYYMMDD&endDate=YYYYMMDD

Two things were confirmed by testing the live page:
  1. No login/cookies are required - it's a public GET endpoint.
  2. The API only accepts windows of at most 30 days, and it rejects any
     startDate earlier than today. So to cover a multi-month range (e.g.
     September through November) the script below breaks the range into
     <=30-day chunks, calls the API once per chunk, and merges/dedupes
     the results.

Usage examples:
    pip install requests openpyxl
    python yangming_schedule_hcm.py
    python yangming_schedule_hcm.py --from 2026-09-01 --to 2026-11-30
    python yangming_schedule_hcm.py --pol CNSHA --pod VNSGN --out schedule.xlsx
"""

import argparse
import datetime as dt

import requests

API_URL = "https://www.yangming.com/api/P2P/GetP2PRoutes"
MAX_WINDOW_DAYS = 30  # the API rejects a wider startDate..endDate span


def daterange_chunks(start: dt.date, end: dt.date, max_days: int = MAX_WINDOW_DAYS):
    """Split [start, end] into consecutive (chunk_start, chunk_end) pairs,
    each spanning at most max_days."""
    cur = start
    while cur <= end:
        chunk_end = min(cur + dt.timedelta(days=max_days), end)
        yield cur, chunk_end
        cur = chunk_end + dt.timedelta(days=1)


def fetch_routes(pol: str, pod: str, start: dt.date, end: dt.date,
                  date_definition: str = "DEP", priority: str = "ALL") -> list:
    """Call Yang Ming's public P2P schedule API for one <=30-day window."""
    params = {
        "locationCodeFrom": pol,
        "serviceTermFrom": "Y",
        "locationCodeTo": pod,
        "serviceTermTo": "Y",
        "priorityWay": priority,          # ALL / DIRECT / TS
        "dateDefinition": date_definition,  # DEP (departure) / ARR (arrival)
        "startDate": start.strftime("%Y%m%d"),
        "endDate": end.strftime("%Y%m%d"),
    }
    resp = requests.get(API_URL, params=params, timeout=30,
                         headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return resp.json()


def collect_schedules(pol: str, pod: str, date_from: dt.date, date_to: dt.date) -> list:
    """Fetch schedules for the whole [date_from, date_to] range, chunked and deduped."""
    today = dt.date.today()
    if date_from < today:
        print(f"Note: start date {date_from} is in the past - Yang Ming's schedule "
              f"API only returns sailings from today ({today}) onward, "
              f"so the start date was moved forward to {today}.")
        date_from = today
    if date_to < date_from:
        raise ValueError("End date must be on or after the (possibly adjusted) start date.")

    rows, seen = [], set()
    for chunk_start, chunk_end in daterange_chunks(date_from, date_to):
        print(f"Fetching {pol} -> {pod}: {chunk_start} to {chunk_end} ...")
        for entry in fetch_routes(pol, pod, chunk_start, chunk_end):
            key = (entry.get("masterVoyageCode"), entry.get("masterETD"), entry.get("masterVesselCode"))
            if key in seen:
                continue
            seen.add(key)
            rows.append(entry)
    return rows


def flatten(entry: dict) -> dict:
    """Turn one API result into a flat dict ready for a spreadsheet row."""
    ts = entry.get("transshipment")
    return {
        "ETD": entry.get("masterETD"),
        "ETA": entry.get("masterETA"),
        "Transit Days": entry.get("transitDays"),
        "Vessel": entry.get("masterVesselName"),
        "Voyage": entry.get("masterComnVoyage"),
        "Voyage Code": entry.get("masterVoyageCode"),
        "POL": entry.get("placeOfReceipt"),
        "POD": entry.get("placeOfDelivery"),
        "Routing": "T/S" if ts else "Direct",
        "Transshipment Port": (ts or {}).get("locationName", "-"),
    }


def export_excel(rows: list, out_path: str):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter

    headers = ["ETD", "ETA", "Transit Days", "Vessel", "Voyage", "Voyage Code",
               "POL", "POD", "Routing", "Transshipment Port"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule"
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    for r in rows:
        ws.append([r[h] for h in headers])

    for i, h in enumerate(headers, start=1):
        width = max(len(h), *(len(str(r[h])) for r in rows)) + 2 if rows else len(h) + 2
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = "A2"
    wb.save(out_path)


def export_csv(rows: list, out_path: str):
    import csv
    headers = ["ETD", "ETA", "Transit Days", "Vessel", "Voyage", "Voyage Code",
               "POL", "POD", "Routing", "Transshipment Port"]
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def main():
    today = dt.date.today()
    parser = argparse.ArgumentParser(
        description="Fetch Yang Ming Point-to-Point sailing schedule and export it.")
    parser.add_argument("--pol", default="CNSHA", help="Origin port code (default: CNSHA = Shanghai)")
    parser.add_argument("--pod", default="VNSGN", help="Destination port code (default: VNSGN = Ho Chi Minh City)")
    parser.add_argument("--from", dest="date_from", default=f"{today.year}-09-01",
                         help="Start date YYYY-MM-DD (default: Sep 1 of the current year)")
    parser.add_argument("--to", dest="date_to", default=f"{today.year}-11-30",
                         help="End date YYYY-MM-DD (default: Nov 30 of the current year)")
    parser.add_argument("--priority", default="ALL", choices=["ALL", "DIRECT", "TS"],
                         help="ALL / DIRECT only / T-S (transshipment) only")
    parser.add_argument("--date-definition", dest="date_definition", default="DEP",
                         choices=["DEP", "ARR"], help="Filter by Departure or Arrival date")
    parser.add_argument("--out", default="yangming_schedule_hcm.xlsx",
                         help="Output file - .xlsx (default) or .csv")
    args = parser.parse_args()

    date_from = dt.date.fromisoformat(args.date_from)
    date_to = dt.date.fromisoformat(args.date_to)

    raw_rows = collect_schedules(args.pol, args.pod, date_from, date_to)
    rows = [flatten(r) for r in raw_rows]
    rows.sort(key=lambda r: (r["ETD"] or "", r["Voyage Code"] or ""))

    print(f"\nFound {len(rows)} sailings from {args.pol} to {args.pod}.")

    if args.out.lower().endswith(".csv"):
        export_csv(rows, args.out)
    else:
        export_excel(rows, args.out)

    print(f"Saved to {args.out}")


if __name__ == "__main__":
    main()
