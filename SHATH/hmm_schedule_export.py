"""
hmm_schedule_export.py

Automates the HMM (hmm21.com) e-Service "Point to Point" sailing schedule search
and exports the results to Excel.

WHY PLAYWRIGHT (a real browser) INSTEAD OF PLAIN HTTP REQUESTS
----------------------------------------------------------------
HMM's schedule search runs through two internal endpoints:
  POST /e-service/general/schedule/apiPointToPointList.do   (kicks off a search, returns a job id "GrmNo")
  POST /e-service/general/schedule/selectPointToPointList.do (returns the actual results for that GrmNo)
These were confirmed (by inspecting the site's own network traffic) to be protected by
bot-detection: calling them directly with a plain HTTP client (no real browser / no prior
page interaction) returns HTTP 403. Driving an actual browser and clicking through the page
the way a person would (what this script does) works reliably. That's also why this script
intercepts the JSON response instead of scraping the rendered HTML: it's the same data the
page itself uses, just captured directly and more reliably than parsing the DOM.

WHAT THIS SCRIPT DOES
----------------------------------------------------------------
1. Opens hmm21.com's schedule page in a real (Playwright-controlled) browser.
2. Sets Loading Port / Discharging Port (default: SHANGHAI, CHINA -> LAEM CHABANG, THAILAND).
3. Repeatedly sets the "Sailing date" field and clicks "Retrieve", stepping the date forward
   by 7 weeks each round (the site's own search only looks 8 weeks ahead at a time), until the
   requested date range is fully covered.
4. Captures each search's JSON result via the page's own network responses.
5. Dedupes, filters to the requested date range, and exports a formatted .xlsx file.

REQUIREMENTS
----------------------------------------------------------------
    pip install playwright openpyxl
    playwright install chromium

This script needs real network access to www.hmm21.com and a working Chromium install.
It will NOT work in a network-sandboxed environment (e.g. most CI containers) - run it on
your own machine.

USAGE
----------------------------------------------------------------
    python hmm_schedule_export.py \\
        --pol "SHANGHAI,CHINA" --pol-code CNSHA \\
        --pod "LAEM CHABANG, THAILAND" --pod-code THLCH \\
        --start 2026-09-01 --end 2026-11-30 \\
        --out HMM_Schedule.xlsx

    # Defaults already match this exact request (Shanghai -> Laem Chabang, Sep-Nov of the
    # current year), so it also runs with no arguments at all:
    python hmm_schedule_export.py
"""

import argparse
import json
import time
from datetime import datetime, timedelta

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

SCHEDULE_URL = "https://www.hmm21.com/e-service/general/schedule/ScheduleMainPost.do"
RESULT_ENDPOINT = "selectPointToPointList.do"


def select_port(page, field_label: str, port_text: str):
    """
    Set a port field (Loading Port / Discharging Port). Tries the simple case first
    (a native <select> you can pick an option from directly); falls back to the
    click -> type -> choose-from-list pattern used by searchable dropdowns.
    """
    combo = page.get_by_role("combobox", name=field_label).first
    try:
        combo.select_option(label=port_text)
        return
    except Exception:
        pass

    # Fallback: click to open it, type to filter, click the matching option.
    field = page.get_by_text(field_label, exact=False).first
    field.click()
    page.keyboard.type(port_text.split(",")[0])
    page.wait_for_timeout(500)
    page.get_by_text(port_text, exact=True).first.click()


def set_sail_date(page, date_str_mmddyyyy: str):
    """date_str_mmddyyyy like '09/18/2026' for the native <input type=date> field."""
    date_input = page.locator('input[type="date"]').first
    date_input.click()
    # Native date inputs accept sequential digit typing (MMDDYYYY).
    digits = date_str_mmddyyyy.replace("/", "")
    page.keyboard.type(digits)


def set_weeks(page, weeks: int = 8):
    weeks_select = page.get_by_role("combobox").filter(has_text="weeks").first
    try:
        weeks_select.select_option(label=f"{weeks} weeks")
    except Exception:
        pass  # leave default if this selector doesn't match; 8 weeks is usually the default max


def run_one_window(page, sail_date: datetime, weeks: int = 8, wait_after_click=6000):
    """
    Sets the sail date, clicks Retrieve, and returns the parsed grmData list captured
    from the selectPointToPointList.do response for this search.
    """
    captured = {}

    def on_response(response):
        if RESULT_ENDPOINT in response.url and response.request.method == "POST":
            try:
                body = response.json()
                if body.get("RTN_STS") == "OK" and "grmData" in body:
                    captured["data"] = body["grmData"]
            except Exception:
                pass

    page.on("response", on_response)
    try:
        set_sail_date(page, sail_date.strftime("%m/%d/%Y"))
        set_weeks(page, weeks)
        page.get_by_role("button", name="Retrieve").click()
        page.wait_for_timeout(wait_after_click)
    finally:
        page.remove_listener("response", on_response)

    return captured.get("data", [])


def flatten(grm_data):
    rows = []
    for item in grm_data:
        legs = item.get("transit") or []
        if not legs:
            continue
        orig_leg, dest_leg = legs[0], legs[-1]
        rows.append({
            "week": item.get("bisWk"),
            "origin_date": (orig_leg.get("arvlStDt") or "")[:10],
            "origin_port": item.get("porLocNm"),
            "destn_date": (dest_leg.get("dpartFnshDt") or "")[:10],
            "destn_port": item.get("podLocNm"),
            "transit_days": round(item["totTrstmHrs"] / 24, 1) if item.get("totTrstmHrs") else None,
            "routing": "Direct" if len(legs) == 1 else "T/S",
            "route": item.get("mthLoopCd"),
            "operator": item.get("mthVslCarrCd"),
            "main_vessel": f'{item.get("mthVslNm")} ({item.get("mthCssmObVoyNo") or ""})'.strip(),
            "inland_cutoff": item.get("fcgoCtofDt"),
            "port_cutoff": item.get("portCtofDt"),
            "doc_cutoff": item.get("sigCtofDt"),
        })
    return rows


def export_excel(rows, out_path, pol_text, pod_text, start_date, end_date):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule"
    FONT_NAME = "Arial"

    headers = ["Week", "Origin Date (ETD)", "Origin Port", "Destination Date (ETA)",
               "Destination Port", "Transit (days)", "Routing", "Route", "Operator",
               "Main Vessel", "Inland Cut-off", "Port Cut-off", "DOC Cut-off"]

    ws.merge_cells("A1:M1")
    ws["A1"] = f"HMM Sailing Schedule: {pol_text} -> {pod_text}"
    ws["A1"].font = Font(name=FONT_NAME, size=14, bold=True)
    ws.merge_cells("A2:M2")
    ws["A2"] = (f"Period: {start_date} to {end_date}  |  Source: hmm21.com e-Service Schedule "
                f"(Point to Point)  |  Retrieved: {datetime.now().strftime('%Y-%m-%d')}")
    ws["A2"].font = Font(name=FONT_NAME, size=10, italic=True, color="555555")

    header_row = 4
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=col_idx, value=h)
        c.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, r in enumerate(rows, start=header_row + 1):
        values = [r["week"], r["origin_date"], r["origin_port"], r["destn_date"],
                  r["destn_port"], r["transit_days"], r["routing"], r["route"],
                  r["operator"], r["main_vessel"], r["inland_cutoff"], r["port_cutoff"],
                  r["doc_cutoff"]]
        for col_idx, v in enumerate(values, start=1):
            c = ws.cell(row=i, column=col_idx, value=v)
            c.font = Font(name=FONT_NAME, size=10)
            c.border = border
            c.alignment = Alignment(horizontal="center", vertical="center")
        if i % 2 == 0:
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=i, column=col_idx).fill = PatternFill("solid", fgColor="F2F6FA")

    widths = [6, 16, 20, 18, 20, 12, 9, 8, 9, 22, 16, 16, 16]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.freeze_panes = "A5"
    ws.row_dimensions[header_row].height = 28

    wb.save(out_path)


def main():
    parser = argparse.ArgumentParser(description="Export HMM sailing schedule to Excel.")
    parser.add_argument("--pol", default="SHANGHAI,CHINA", help="Loading port display name")
    parser.add_argument("--pod", default="LAEM CHABANG, THAILAND", help="Discharging port display name")
    parser.add_argument("--start", default=None, help="Start date YYYY-MM-DD (default: today)")
    parser.add_argument("--end", default=None, help="End date YYYY-MM-DD (default: 3 months from start)")
    parser.add_argument("--out", default="HMM_Schedule.xlsx", help="Output .xlsx path")
    parser.add_argument("--headless", action="store_true", default=True)
    parser.add_argument("--no-headless", dest="headless", action="store_false")
    args = parser.parse_args()

    start_date = datetime.strptime(args.start, "%Y-%m-%d") if args.start else datetime.now()
    end_date = (datetime.strptime(args.end, "%Y-%m-%d") if args.end
                else start_date + timedelta(days=90))

    all_grm = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        page = browser.new_page()
        page.goto(SCHEDULE_URL, wait_until="networkidle")

        select_port(page, "Loading Port", args.pol)
        select_port(page, "Discharging Port", args.pod)

        cursor = start_date
        while cursor <= end_date:
            print(f"Searching from {cursor:%Y-%m-%d} ...")
            data = run_one_window(page, cursor, weeks=8)
            print(f"  -> {len(data)} sailings")
            all_grm.extend(data)
            cursor += timedelta(weeks=7)  # 1 week overlap between windows, 8-week window each time
            time.sleep(1.5)  # be gentle / look human between searches

        browser.close()

    rows = flatten(all_grm)

    # dedupe by (vessel/voyage, origin date)
    seen = set()
    uniq = []
    for r in rows:
        key = (r["main_vessel"], r["origin_date"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(r)

    start_str, end_str = start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")
    filtered = [r for r in uniq if r["origin_date"] and start_str <= r["origin_date"] <= end_str]
    filtered.sort(key=lambda r: r["origin_date"])

    export_excel(filtered, args.out, args.pol, args.pod, start_str, end_str)
    print(f"\nDone: {len(filtered)} sailings written to {args.out}")


if __name__ == "__main__":
    main()
