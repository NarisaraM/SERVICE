#!/usr/bin/env python3
"""
ดึงตารางเรือ MSC (Point-to-Point) จาก Shanghai -> Manzanillo (Mexico)
ช่วงเดือน กันยายน - พฤศจิกายน แล้วบันทึกเป็นไฟล์ Excel/CSV

API ที่หน้าเว็บ https://www.msc.com/en/search-a-schedule เรียกใช้จริง:
    POST https://www.msc.com/api/feature/tools/SearchSailingRoutes
    body: {"FromDate": "YYYY-MM-DD", "fromPortId": 444, "toPortId": 236,
           "language": "en", "dataSourceId": "{E9CCBD25-...}"}
    header ที่จำเป็น: X-Requested-With: XMLHttpRequest
    (fromPortId 444 = SHANGHAI, toPortId 236 = MANZANILLO, MEXICO, MXZLO)
    หมายเหตุ: มี Manzanillo อีกแห่งในปานามา (id 1111) แต่ไม่มีตารางจาก Shanghai

ข้อจำกัดของ API: FromDate ห้ามเป็นวันที่ผ่านมาแล้ว ("Date cannot be in the past")
สคริปต์จึงเริ่มค้นจากวันนี้ ถ้าเดือนเริ่มต้นผ่านไปแล้ว

ติดตั้ง:  pip install requests pandas openpyxl
รัน:      python msc_schedule_shanghai_manzanillo.py
          python msc_schedule_shanghai_manzanillo.py --year 2026 --start-month 9 --end-month 11
"""
import argparse
import calendar
import sys
import time
from datetime import date, datetime

import pandas as pd
import requests

BASE = "https://www.msc.com"
PAGE_URL = f"{BASE}/en/search-a-schedule"
API_URL = f"{BASE}/api/feature/tools/SearchSailingRoutes"

FROM_PORT_ID = 444   # SHANGHAI (CNSHA)
TO_PORT_ID = 236     # MANZANILLO, MEXICO (MXZLO)
DATA_SOURCE_ID = "{E9CCBD25-6FBA-4C5C-85F6-FC4F9E5A931F}"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Content-Type": "application/json",
    "X-Requested-With": "XMLHttpRequest",   # จำเป็น ไม่งั้น API ตอบกลับว่าง
    "Origin": BASE,
    "Referer": PAGE_URL,
}


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    s.get(PAGE_URL, timeout=30)  # เปิดหน้าเว็บก่อนเพื่อรับ cookie
    return s


def search(session: requests.Session, from_date: date) -> list[dict]:
    """เรียก API หนึ่งครั้ง คืนรายการ route (แต่ละเที่ยวเรือ)"""
    payload = {
        "FromDate": from_date.isoformat(),
        "fromPortId": FROM_PORT_ID,
        "toPortId": TO_PORT_ID,
        "language": "en",
        "dataSourceId": DATA_SOURCE_ID,
    }
    r = session.post(API_URL, json=payload, timeout=30)
    r.raise_for_status()
    if not r.text.strip() or r.text.strip() == '""':
        raise RuntimeError("API ตอบกลับว่าง (อาจถูกบล็อก/ขาด header) - ดูหมายเหตุท้ายไฟล์")
    j = r.json()
    data = j.get("Data")
    if not j.get("IsSuccess") or not isinstance(data, list):
        # เช่น "There are currently no results available for the selected port pair."
        return []

    rows = []
    for group in data:
        service = group.get("LoadingService") or group["Key"].get("MaritimeServiceName")
        for rt in group.get("Routes", []):
            cut = rt.get("CutOffs") or {}
            rows.append({
                "Service": service,
                "Vessel": rt.get("VesselName"),
                "Voyage": rt.get("DepartureVoyageNo"),
                "POL": group.get("PortOfLoad"),
                "POD": group.get("PortOfDischarge"),
                "ETD": rt.get("EstimatedDepartureDate"),
                "ETA": rt.get("EstimatedArrivalDate"),
                "TransitTime": rt.get("TotalTransitTime"),
                "RoutingType": group.get("RoutingType"),   # Direct / Transhipment
                "Legs": len(rt.get("RouteScheduleLegDetails") or []),
                "CO2": rt.get("CO2FootPrint"),
                "CY CutOff": cut.get("ContainerYardCutOffDate"),
                "Reefer CutOff": cut.get("ReeferCutOffDate"),
                "DG CutOff": cut.get("DangerousCargoCutOffDate"),
                "SI CutOff": cut.get("ShippingInstructionsCutOffDate"),
                "VGM CutOff": cut.get("VerifiedGrossMassCutOffDate"),
            })
    return rows


def collect(year: int, m_start: int, m_end: int) -> pd.DataFrame:
    win_start = date(year, m_start, 1)
    win_end = date(year, m_end, calendar.monthrange(year, m_end)[1])
    today = date.today()
    if win_end < today:
        sys.exit(f"ช่วงที่เลือก ({win_start} - {win_end}) ผ่านไปแล้ว API ไม่รองรับวันที่ในอดีต")

    session = make_session()
    seen, all_rows = set(), []
    cursor = max(win_start, today)

    # API คืนเที่ยวเรือตั้งแต่ FromDate เป็นต้นไปหลายสัปดาห์ จึงเลื่อนวันที่ค้นต่อไปเรื่อย ๆ จนพ้นช่วง
    while cursor <= win_end:
        rows = search(session, cursor)
        if not rows:
            break
        new_last = None
        for row in rows:
            key = (row["Service"], row["Vessel"], row["Voyage"], row["ETD"])
            etd = datetime.fromisoformat(row["ETD"]).date()
            new_last = etd if new_last is None or etd > new_last else new_last
            if key not in seen:
                seen.add(key)
                all_rows.append(row)
        if new_last is None or new_last <= cursor:
            break
        cursor = new_last  # ค้นต่อจากเที่ยวสุดท้ายที่ได้
        time.sleep(1)      # ไม่ยิง API ถี่เกินไป
        if len(rows) < 2:
            break

    df = pd.DataFrame(all_rows)
    if df.empty:
        return df
    df["ETD"] = pd.to_datetime(df["ETD"])
    df["ETA"] = pd.to_datetime(df["ETA"])
    mask = (df["ETD"].dt.date >= win_start) & (df["ETD"].dt.date <= win_end)
    return df[mask].sort_values("ETD").reset_index(drop=True)


def main():
    ap = argparse.ArgumentParser(description="MSC schedule Shanghai -> Manzanillo (Mexico)")
    ap.add_argument("--year", type=int, default=date.today().year)
    ap.add_argument("--start-month", type=int, default=9)
    ap.add_argument("--end-month", type=int, default=11)
    ap.add_argument("--out", default="MSC_Shanghai_Manzanillo")
    args = ap.parse_args()

    df = collect(args.year, args.start_month, args.end_month)
    if df.empty:
        print("ไม่พบตารางเรือในช่วงที่เลือก")
        return

    print(df[["Service", "RoutingType", "Vessel", "Voyage", "ETD", "ETA", "TransitTime"]].to_string(index=False))
    df.to_excel(f"{args.out}.xlsx", index=False)
    df.to_csv(f"{args.out}.csv", index=False, encoding="utf-8-sig")
    print(f"\nพบ {len(df)} เที่ยว -> บันทึก {args.out}.xlsx / .csv แล้ว")


if __name__ == "__main__":
    main()

# หมายเหตุ: ถ้าเรียกจากสคริปต์แล้วได้ผลว่าง/403 (เว็บ MSC มีระบบกันบอท)
# ให้ลองรันจากเครื่องที่เปิดเว็บ MSC ได้ปกติ หรือเปลี่ยนไปใช้ Playwright เปิดหน้า
# https://www.msc.com/en/search-a-schedule แล้วเรียก fetch() ภายในหน้าแทน requests
