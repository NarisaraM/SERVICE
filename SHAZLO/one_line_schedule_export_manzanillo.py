"""
one_line_schedule_export_manzanillo.py
======================================

ดึงตารางเรือ (Point to Point Schedule) จากเว็บไซต์ ONE (Ocean Network Express)
https://www.one-line.com/one-ecom/schedule/point-to-point-schedule
โดย automate การค้นหาด้วย Playwright แล้วกดปุ่ม "Download" -> "Excel File"
ของหน้าเว็บจริง (ใช้ไฟล์ที่เว็บ ONE สร้างให้ ไม่ได้ scrape ตาราง HTML เอง)
จากนั้นรวมไฟล์ทั้งหมดเป็น Excel ไฟล์เดียว

ค่าเริ่มต้นของสคริปต์นี้ถูกตั้งไว้ตามคำขอ:
    Origin (POL)      : SHANGHAI, SHANGHAI, CHINA   (CNSHA)
    Destination (POD) : MANZANILLO, MEXICO         (MXZLO)
    ช่วงวันที่ (ETD)   : 2026-09-01 ถึง 2026-11-30 (กันยายน - พฤศจิกายน 2026)

รหัสท่าเรือ CNSHA / MXZLO และชื่อท่าเรือด้านบน ตรวจสอบมาจาก URL จริงที่หน้าเว็บ
ONE สร้างขึ้นเวลากดค้นหา Shanghai -> Manzanillo บนหน้าเว็บ ไม่ได้เดาเอง

เหตุผลที่ต้อง "แบ่งช่วง" การค้นหา
----------------------------------
หน้าเว็บ ONE จำกัดผลลัพธ์การค้นหาต่อครั้งไว้ที่ "Next 8 Weeks" (56 วัน) เท่านั้น
(ตัวเลือกในหน้าเว็บมีแค่ 2 / 4 / 6 / 8 สัปดาห์) สคริปต์นี้จึงแบ่งช่วง 3 เดือนที่ขอ
ออกเป็นหลาย "หน้าต่าง" (window) ละไม่เกิน 56 วัน วนดาวน์โหลดทีละหน้าต่าง แล้วค่อย
นำมารวมกันในภายหลัง (มีการตัดแถวที่ซ้ำกันจากช่วงคาบเกี่ยวออกให้)

วิธีใช้งาน
----------
1) ติดตั้งไลบรารีที่ต้องใช้ (รันครั้งเดียว):

    pip install playwright pandas openpyxl
    playwright install chromium

2) รันสคริปต์ (ใช้ค่าเริ่มต้น Shanghai -> Manzanillo, ก.ย.-พ.ย. 2026):

    python one_line_schedule_export_manzanillo.py

   หรือกำหนดพารามิเตอร์เอง เช่นเปลี่ยนเส้นทาง/ช่วงวันที่:

    python one_line_schedule_export_manzanillo.py \
        --origin-code CNSHA --origin-name "SHANGHAI, SHANGHAI, CHINA" \
        --dest-code MXZLO   --dest-name "MANZANILLO, MEXICO" \
        --start 2026-09-01 --end 2026-11-30 \
        --output ONE_Shanghai_HoChiMinh_Sep-Nov2026.xlsx

หมายเหตุ
--------
- สคริปต์นี้ต้อง "รันในเครื่อง/เซิร์ฟเวอร์ที่ต่ออินเทอร์เน็ตออกไปหา www.one-line.com ได้ปกติ"
  (รันไม่ได้ในแซนด์บ็อกซ์ที่ถูกบล็อก outbound network)
- เว็บไซต์อาจมีการเปลี่ยนโครงสร้างหน้า (selector) ได้ในอนาคต หากสคริปต์ error ที่ขั้นตอน
  รอผลลัพธ์ (wait_for_selector) หรือกดปุ่ม Download ให้ลองเปิดหน้าเว็บด้วยตา (headless=False)
  เพื่อดูว่าหน้าตาต่างไปจากเดิมหรือไม่
- โปรดตรวจสอบเงื่อนไขการใช้งาน (Terms of Use) ของเว็บไซต์ ONE ก่อนใช้งานอัตโนมัติในเชิงพาณิชย์/ถี่ๆ
"""

from __future__ import annotations

import argparse
import asyncio
import sys
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import quote

import pandas as pd

try:
    from playwright.async_api import async_playwright, TimeoutError as PWTimeoutError
except ImportError:
    print(
        "ไม่พบไลบรารี playwright กรุณาติดตั้งก่อน:\n"
        "    pip install playwright pandas openpyxl\n"
        "    playwright install chromium",
        file=sys.stderr,
    )
    raise

BASE_URL = "https://www.one-line.com/one-ecom/schedule/point-to-point-schedule"
MAX_WINDOW_DAYS = 56  # เว็บ ONE จำกัดไว้ที่ "8 Weeks" ต่อการค้นหา 1 ครั้ง


@dataclass
class SearchWindow:
    start: date
    days: int


def build_search_url(
    window: SearchWindow,
    origin_code: str,
    origin_name: str,
    dest_code: str,
    dest_name: str,
    cargo_nature: str = "GP",
) -> str:
    """สร้าง URL ของหน้า Point to Point Schedule พร้อมพารามิเตอร์ค้นหา
    (พารามิเตอร์เหล่านี้ตรวจสอบมาจาก URL จริงที่หน้าเว็บ ONE ใช้เวลากดค้นหา)
    """
    params = {
        "oriLocNmParam": origin_name,
        "destLocNmParam": dest_name,
        "oriLocCdParam": origin_code,
        "destLocCdParam": dest_code,
        "oriTermCdPara": "Y",
        "desTermCdPara": "Y",
        "frmDtParam": window.start.isoformat(),
        "nextWeekValue": str(window.days),
        "cargoNature": cargo_nature,
        "isEnabledCO2": "false",
        "year": str(window.start.year),
        "month": str(window.start.month),
        "searchType": "List",
        "isPolPodOn": "false",
    }
    query = "&".join(f"{k}={quote(str(v))}" for k, v in params.items())
    return f"{BASE_URL}?{query}"


def build_windows(start: date, end: date, step_days: int = MAX_WINDOW_DAYS) -> list[SearchWindow]:
    """แบ่งช่วง [start, end] ออกเป็นหน้าต่างละไม่เกิน step_days วัน (มีคาบเกี่ยวกันได้เล็กน้อย
    เพื่อให้แน่ใจว่าครอบคลุมทุกวันจนถึง end)"""
    windows: list[SearchWindow] = []
    cur = start
    while cur <= end:
        windows.append(SearchWindow(start=cur, days=step_days))
        cur += timedelta(days=step_days)
    return windows


async def dismiss_cookie_banner(page) -> None:
    """พยายามปิด cookie/consent banner ถ้ามี (ไม่ error ถ้าไม่เจอ)"""
    for text in ["Accept", "ACCEPT", "I Agree", "Accept All"]:
        try:
            btn = page.get_by_role("button", name=text, exact=False)
            if await btn.count():
                await btn.first.click(timeout=3000)
                return
        except Exception:
            pass


async def download_window(
    page,
    window: SearchWindow,
    origin_code: str,
    origin_name: str,
    dest_code: str,
    dest_name: str,
    cargo_nature: str,
    download_dir: Path,
) -> Path | None:
    url = build_search_url(window, origin_code, origin_name, dest_code, dest_name, cargo_nature)
    print(f"[{window.start.isoformat()}] เปิดหน้าเว็บ: {url}")
    await page.goto(url, wait_until="domcontentloaded", timeout=60000)
    await dismiss_cookie_banner(page)

    try:
        await page.wait_for_selector(r"text=/Total\s+\d+\s+results/", timeout=30000)
    except PWTimeoutError:
        print(f"  [!] ไม่พบผลลัพธ์ในช่วงนี้ (ข้าม): {window.start.isoformat()}")
        return None

    total_text = await page.locator(r"text=/Total\s+\d+\s+results/").first.text_content()
    print(f"  ผลลัพธ์: {total_text.strip() if total_text else 'N/A'}")

    # ปุ่ม "Download" ตัวแรกบนหน้า คือปุ่มเปิด modal เลือกไฟล์ (ไม่ใช่ปุ่มในแถวผลลัพธ์)
    download_button = page.locator("button:has-text('Download')").first
    await download_button.click()

    dialog = page.locator("[role=dialog]")
    await dialog.wait_for(state="visible", timeout=10000)

    excel_radio = dialog.get_by_role("radio", name="Excel File")
    if await excel_radio.count():
        await excel_radio.check()

    out_path = download_dir / f"one_schedule_{window.start.isoformat()}.xlsx"
    try:
        async with page.expect_download(timeout=30000) as dl_info:
            await dialog.get_by_role("button", name="Download", exact=True).click()
        download = await dl_info.value
        await download.save_as(out_path)
        print(f"  บันทึกไฟล์: {out_path}")
        return out_path
    except PWTimeoutError:
        print(f"  [!] ดาวน์โหลดไม่สำเร็จสำหรับช่วง {window.start.isoformat()}")
        return None


def merge_excel_files(files: list[Path], start: date, end: date, output_path: Path) -> pd.DataFrame:
    """รวมไฟล์ Excel ที่ดาวน์โหลดมาทั้งหมดเป็นไฟล์เดียว ตัดแถวที่ซ้ำกัน (จากช่วงคาบเกี่ยว)
    และพยายามกรองตามช่วงวันที่ที่ขอ ถ้าหาคอลัมน์วันที่ ETD/Departure เจอ"""
    frames = []
    for f in files:
        try:
            df = pd.read_excel(f)
            df["__source_file"] = f.name
            frames.append(df)
        except Exception as exc:
            print(f"  [!] อ่านไฟล์ {f} ไม่สำเร็จ: {exc}")

    if not frames:
        raise RuntimeError("ไม่มีไฟล์ Excel ที่อ่านได้สำเร็จเลย ไม่สามารถรวมผลลัพธ์ได้")

    combined = pd.concat(frames, ignore_index=True)

    # ตัดแถวซ้ำ (ไม่รวมคอลัมน์ __source_file ตอนเทียบ)
    compare_cols = [c for c in combined.columns if c != "__source_file"]
    combined = combined.drop_duplicates(subset=compare_cols).reset_index(drop=True)

    # พยายามหาคอลัมน์วันที่ออกเรือ (ETD) เพื่อกรองตามช่วงที่ขอ และเรียงลำดับ
    date_col_candidates = [
        c for c in combined.columns
        if any(k in str(c).upper() for k in ("ETD", "DEPARTURE", "DEPART"))
    ]
    if date_col_candidates:
        date_col = date_col_candidates[0]
        parsed = pd.to_datetime(combined[date_col], errors="coerce")
        mask = parsed.dt.date.between(start, end)
        # กรองเฉพาะเมื่อยังเหลือข้อมูลอยู่ (กันกรณี parse ผิดจนข้อมูลหายหมด)
        if mask.any():
            combined = combined.loc[mask].copy()
            combined["_sort_date"] = parsed.loc[mask]
            combined = combined.sort_values("_sort_date").drop(columns="_sort_date")
        else:
            print(f"  [!] กรองตามคอลัมน์ '{date_col}' แล้วไม่เหลือข้อมูล จึงใช้ข้อมูลทั้งหมดโดยไม่กรอง")
    else:
        print("  [!] ไม่พบคอลัมน์วันที่ ETD/Departure ที่ชัดเจน จึงไม่ได้กรองตามช่วงวันที่ (รวมทุกแถวที่ดาวน์โหลดมา)")

    combined.to_excel(output_path, index=False)
    return combined


async def run(args: argparse.Namespace) -> None:
    start = date.fromisoformat(args.start)
    end = date.fromisoformat(args.end)
    windows = build_windows(start, end)

    download_dir = Path(args.download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)

    downloaded: list[Path] = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=args.headless)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        for window in windows:
            path = await download_window(
                page,
                window,
                args.origin_code,
                args.origin_name,
                args.dest_code,
                args.dest_name,
                args.cargo,
                download_dir,
            )
            if path:
                downloaded.append(path)

        await browser.close()

    if not downloaded:
        print("ไม่สามารถดาวน์โหลดไฟล์ตารางเรือได้เลย กรุณาตรวจสอบการเชื่อมต่อ/โครงสร้างหน้าเว็บ")
        sys.exit(1)

    print(f"\nรวมไฟล์ทั้งหมด {len(downloaded)} ไฟล์ -> {args.output}")
    combined = merge_excel_files(downloaded, start, end, Path(args.output))
    print(f"เสร็จสิ้น: {len(combined)} แถว, บันทึกเป็น {args.output}")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="ดึงตารางเรือจากเว็บ ONE แล้ว export เป็น Excel")
    parser.add_argument("--origin-code", default="CNSHA", help="รหัสท่าเรือต้นทาง (UN/LOCODE)")
    parser.add_argument("--origin-name", default="SHANGHAI, SHANGHAI, CHINA", help="ชื่อท่าเรือต้นทางตามที่เว็บ ONE ใช้")
    parser.add_argument("--dest-code", default="MXZLO", help="รหัสท่าเรือปลายทาง (UN/LOCODE)")
    parser.add_argument("--dest-name", default="MANZANILLO, MEXICO", help="ชื่อท่าเรือปลายทางตามที่เว็บ ONE ใช้")
    parser.add_argument("--start", default="2026-09-01", help="วันที่เริ่มต้น (YYYY-MM-DD)")
    parser.add_argument("--end", default="2026-11-30", help="วันที่สิ้นสุด (YYYY-MM-DD)")
    parser.add_argument("--cargo", default="GP", help="ประเภทตู้สินค้า (GP = Dry/General)")
    parser.add_argument("--output", default="ONE_Schedule_Export.xlsx", help="ชื่อไฟล์ Excel ผลลัพธ์ที่รวมแล้ว")
    parser.add_argument("--download-dir", default="downloads", help="โฟลเดอร์เก็บไฟล์ Excel รายช่วงก่อนรวม")
    parser.add_argument("--headless", action="store_true", default=True, help="รันแบบไม่เปิดหน้าต่างเบราว์เซอร์ (ค่าเริ่มต้น)")
    parser.add_argument("--no-headless", dest="headless", action="store_false", help="เปิดหน้าต่างเบราว์เซอร์ให้ดู (debug)")
    return parser.parse_args(argv)


if __name__ == "__main__":
    asyncio.run(run(parse_args()))
