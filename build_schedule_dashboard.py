#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_schedule_dashboard.py
============================

Orchestrator that reads/runs the per-carrier schedule scripts found in the
SHATH (Shanghai -> Laem Chabang), SHASGN (Shanghai -> Ho Chi Minh) and
SHAZLO (Shanghai -> Manzanillo) folders next to this file, merges every
carrier's sailings into one dataset, collapses sailings that are really the
same physical vessel call shared across carriers (slot-sharing / alliance
sailings) into a single calendar entry that lists every line selling that
sailing, and renders an interactive HTML calendar/dashboard
(dashboard/index.html) with a route selector and clickable carrier-logo
filters.

Carrier sources
---------------
  Live (fetched every run, cached in <route>/route_live_rows.json):
    COSCO      elines.coscoshipping.com public JSON API
    Yang Ming  yangming.com public JSON API
    ONE        one-line.com point-to-point page, read with Playwright
  Files dropped into each route folder (re-read on every run):
    HMM        HMM*.xls (HTML table download) or HMM_*.xlsx
    Maersk     Maersk*.csv (or maersk_*.xlsx)
    OOCL       Oocl*.xlsx
    MSC        MSC_*.xlsx   (msc.com blocks automated access, so no live fetch)
    Hapag-Lloyd HapagLloyd_*.xlsx
    ONE        ONE_Schedule_*.xlsx (routing inferred from transit time: > 11 days = T/S)

Usage
-----
    python build_schedule_dashboard.py                 # fetch + rebuild everything
    python build_schedule_dashboard.py --use-cache      # reuse the cached live COSCO/Yang Ming/ONE rows (files are always re-read)
    python build_schedule_dashboard.py --skip-one       # skip the (slower) ONE scrape
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import shutil
import sys
from pathlib import Path
from urllib.parse import quote

import requests

try:
    import pandas as pd
except ImportError:
    pd = None

BASE_DIR = Path(__file__).resolve().parent
DASHBOARD_DIR = BASE_DIR / "dashboard"
ASSETS_DIR = DASHBOARD_DIR / "assets"

TODAY = dt.date.today()
DEFAULT_START = dt.date(TODAY.year, 9, 1)
DEFAULT_END = dt.date(TODAY.year, 11, 30)

# ---------------------------------------------------------------------------
# Route configuration - one entry per Shanghai -> X lane / folder
# ---------------------------------------------------------------------------
ROUTES = [
    {
        "code": "SHATH",
        "folder": BASE_DIR / "SHATH",
        "name": "Laem Chabang, Thailand",
        "short": "Laem Chabang",
        "cosco_dest_uuid": "738872886233472",
        "cosco_dest_city": "Laem Chabang, ,Chon Buri,Thailand,THLCB",
        "yml_pod": "THLCB",
        "one_dest_code": "THLCH",
        "one_dest_name": "LAEM CHABANG, THAILAND",
        "msc_port_id": 36,
        "hmm_dest": "LAEM CHABANG",
        "hmm_dest_code": "THLCH",
    },
    {
        "code": "SHASGN",
        "folder": BASE_DIR / "SHASGN",
        "name": "Ho Chi Minh, Vietnam",
        "short": "Ho Chi Minh",
        "cosco_dest_uuid": "882736036768074",
        "cosco_dest_city": "Ho Chi Minh (Cat Lai), ,Ho Chi Minh,Vietnam,VNCAL",
        "yml_pod": "VNSGN",
        "one_dest_code": "VNSGN",
        "one_dest_name": "HO CHI MINH, VIETNAM",
        "msc_port_id": 52,
        "hmm_dest": "HOCHIMINH",
        "hmm_dest_code": "VNSGN",
    },
    {
        "code": "SHAZLO",
        "folder": BASE_DIR / "SHAZLO",
        "name": "Manzanillo, Mexico",
        "short": "Manzanillo",
        "cosco_dest_uuid": "738872886247648",
        "cosco_dest_city": "Manzanillo, ,Colima,Mexico,MXZLO",
        "yml_pod": "MXZLO",
        "one_dest_code": "MXZLO",
        "one_dest_name": "MANZANILLO, MEXICO",
        "msc_port_id": 236,
        "hmm_dest": "MANZANILLO",
        "hmm_dest_code": "MXZLO",
    },
]

CARRIERS = {
    "COSCO": {"label": "COSCO Shipping", "color": "#1f4fa8", "logo": "cosco.png"},
    "ONE": {"label": "ONE", "color": "#e2007a", "logo": "one.webp"},
    "HMM": {"label": "HMM", "color": "#f28c00", "logo": "hmm.webp"},
    "YML": {"label": "Yang Ming", "color": "#1a8a4a", "logo": "yangming.png"},
    "MAERSK": {"label": "Maersk", "color": "#2aa3d6", "logo": "maersk.png"},
    "OOCL": {"label": "OOCL", "color": "#e5352b", "logo": "oocl.png"},
    "MSC": {"label": "MSC", "color": "#c9a227", "logo": "msc.jpg"},
    "HAPAG": {"label": "Hapag-Lloyd", "color": "#f26b21", "logo": "hapag.png"},
}

ORIGIN_CITY_UUID = "738872886232873"
ORIGIN_CITY = "Shanghai,Shanghai,Shanghai,China,CNSHA"


# ---------------------------------------------------------------------------
# COSCO - direct JSON API (see SHATH/cosco_schedule_shanghai_laemchabang.py)
# ---------------------------------------------------------------------------
def fetch_cosco(route: dict, start: dt.date, end: dt.date) -> list[dict]:
    url = "https://elines.coscoshipping.com/ebschedule/public/purpoShipmentWs"
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"),
        "Origin": "https://elines.coscoshipping.com",
        "Referer": "https://elines.coscoshipping.com/ebusiness/sailingSchedule/searchByCity/resultByCity",
    }
    payload = {
        "fromDate": start.isoformat(), "toDate": end.isoformat(),
        "pickup": "B", "delivery": "B", "estimateDate": "D",
        "originCityUuid": ORIGIN_CITY_UUID, "destinationCityUuid": route["cosco_dest_uuid"],
        "originCity": ORIGIN_CITY, "destinationCity": route["cosco_dest_city"],
        "cargoNature": "",
    }
    resp = requests.post(url, json=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    body = resp.json()
    if str(body.get("code")) != "200":
        raise RuntimeError(f"COSCO API error: {body}")
    raw = body.get("data", {}).get("content", {}).get("data", []) or []

    rows = []
    for r in raw:
        rows.append({
            "carrier": "COSCO",
            "vessel": (r.get("vessel") or "").strip(),
            "voyage": (r.get("extVoyage") or "").strip(),
            "service": r.get("service"),
            "pol": r.get("pol") or "Shanghai",
            "etd": (r.get("etd") or "")[:10],
            "pod": r.get("pod") or route["short"],
            "eta": (r.get("eta") or "")[:10],
            "transit_days": r.get("transitTime"),
            "routing": "Direct",
            "cutoff": r.get("cutOff"),
        })
    return rows


# ---------------------------------------------------------------------------
# Yang Ming - direct JSON API (see SHATH/yangming_schedule.py)
# ---------------------------------------------------------------------------
def fetch_yangming(route: dict, start: dt.date, end: dt.date) -> list[dict]:
    url = "https://www.yangming.com/api/P2P/GetP2PRoutes"
    if start < TODAY:
        start = TODAY
    if end < start:
        return []

    rows, seen = [], set()
    cur = start
    while cur <= end:
        chunk_end = min(cur + dt.timedelta(days=30), end)
        params = {
            "locationCodeFrom": "CNSHA", "serviceTermFrom": "Y",
            "locationCodeTo": route["yml_pod"], "serviceTermTo": "Y",
            "priorityWay": "ALL", "dateDefinition": "DEP",
            "startDate": cur.strftime("%Y%m%d"), "endDate": chunk_end.strftime("%Y%m%d"),
        }
        resp = requests.get(url, params=params, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
        for entry in resp.json():
            key = (entry.get("masterVoyageCode"), entry.get("masterETD"), entry.get("masterVesselCode"))
            if key in seen:
                continue
            seen.add(key)
            ts = entry.get("transshipment")
            rows.append({
                "carrier": "YML",
                "vessel": (entry.get("masterVesselName") or "").strip(),
                "voyage": entry.get("masterComnVoyage") or entry.get("masterVoyageCode"),
                "service": None,
                "pol": entry.get("placeOfReceipt") or "Shanghai",
                "etd": (entry.get("masterETD") or "")[:10],
                "pod": entry.get("placeOfDelivery") or route["short"],
                "eta": (entry.get("masterETA") or "")[:10],
                "transit_days": entry.get("transitDays"),
                "routing": "T/S" if ts else "Direct",
                "cutoff": None,
            })
        cur = chunk_end + dt.timedelta(days=1)
    return rows


# ---------------------------------------------------------------------------
# ONE - Playwright scrape of the rendered result cards (see scratchpad
# one_scraper.py this was developed from: the site's Download-button flow is
# blocked by a cookie-consent overlay, so we read the cards' own text instead)
# ---------------------------------------------------------------------------
ONE_BASE_URL = "https://www.one-line.com/one-ecom/schedule/point-to-point-schedule"
ONE_CARD_RE = re.compile(
    r"Origin\n+(?P<etd>\d{4}-\d{2}-\d{2})\n+(?P<pol>[^\n]+)\n+"
    r"(?P<transit>\d+) day\(s\)\n+(?P<routing>Direct|Transshipment)\n+"
    r"(?P<pod>[^\n]+)\n+Destination\n+(?P<eta>\d{4}-\d{2}-\d{2})\n+"
    r"Vessel Voyage / Service Lane\n+(?P<vessel_voyage>[^\n]+)\n+/\n+(?P<service_lane>[^\n]+)\n+"
    r"Inland Cut-off\n+(?P<inland_cutoff>[^\n]+)\n+Port Cut-off\n+(?P<port_cutoff>[^\n]+)"
)


def _one_build_url(dest_code: str, dest_name: str, start: dt.date, days: int) -> str:
    params = {
        "oriLocNmParam": "SHANGHAI, SHANGHAI, CHINA", "destLocNmParam": dest_name,
        "oriLocCdParam": "CNSHA", "destLocCdParam": dest_code,
        "oriTermCdPara": "Y", "desTermCdPara": "Y",
        "frmDtParam": start.isoformat(), "nextWeekValue": str(days),
        "cargoNature": "GP", "isEnabledCO2": "false",
        "year": str(start.year), "month": str(start.month),
        "searchType": "List", "isPolPodOn": "false",
    }
    query = "&".join(f"{k}={quote(str(v))}" for k, v in params.items())
    return f"{ONE_BASE_URL}?{query}"


def _one_split_vessel_voyage(vv: str) -> tuple[str, str]:
    m = re.match(r"^(.*\D)\s*(\d[\dA-Z]*[EWSN])$", vv.strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return vv.strip(), ""


def fetch_one(route: dict, start: dt.date, end: dt.date, max_pages: int = 35) -> list[dict]:
    from playwright.sync_api import sync_playwright

    windows = []
    cur = start
    while cur <= end:
        windows.append(cur)
        cur += dt.timedelta(days=49)  # 56-day window, 7-day overlap

    all_rows, seen = [], set()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        for w in windows:
            url = _one_build_url(route["one_dest_code"], route["one_dest_name"], w, 56)
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(2500)
            for sel in ["button:has-text('Accept All')", "text=Accept All"]:
                try:
                    btn = page.locator(sel).first
                    if btn.count() and btn.is_visible(timeout=1000):
                        btn.click(timeout=2000)
                        break
                except Exception:
                    pass
            try:
                page.wait_for_selector("text=Vessel Voyage / Service Lane", timeout=40000)
            except Exception:
                continue
            page.wait_for_timeout(1200)

            for pg in range(max_pages):
                text = page.inner_text("body")
                for m in ONE_CARD_RE.finditer(text):
                    d = m.groupdict()
                    key = (d["vessel_voyage"], d["etd"])
                    if key in seen:
                        continue
                    seen.add(key)
                    vessel, voyage = _one_split_vessel_voyage(d["vessel_voyage"])
                    all_rows.append({
                        "carrier": "ONE",
                        "vessel": vessel,
                        "voyage": voyage,
                        "service": d["service_lane"],
                        "pol": d["pol"],
                        "etd": d["etd"],
                        "pod": d["pod"],
                        "eta": d["eta"],
                        "transit_days": int(d["transit"]),
                        "routing": "Direct" if d["routing"] == "Direct" else "T/S",
                        "cutoff": d["port_cutoff"],
                    })
                next_btn = page.locator("text=Next Page").first
                try:
                    if next_btn.count() and next_btn.is_enabled(timeout=1000):
                        next_btn.click(timeout=2000)
                        page.wait_for_timeout(1000)
                    else:
                        break
                except Exception:
                    break
        browser.close()
    return [r for r in all_rows if start.isoformat() <= r["etd"] <= end.isoformat()]


# ---------------------------------------------------------------------------
# HMM - best-effort live scrape, otherwise fall back to a cached export
# ---------------------------------------------------------------------------
def fetch_hmm_live(route: dict, start: dt.date, end: dt.date) -> list[dict]:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

    url = "https://www.hmm21.com/e-service/general/schedule/ScheduleMainPost.do"
    rows = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url, wait_until="load", timeout=25000)  # fails fast if the site is unreachable
        page.locator("#srchPointFrom").click()
        page.locator("#srchPointFrom").fill("SHANGHAI")
        page.wait_for_selector("div.ac_results li", timeout=8000)
        page.locator("div.ac_results li").first.click()
        page.locator("#srchPointTo").click()
        page.locator("#srchPointTo").fill(route["hmm_dest"])
        page.wait_for_selector("div.ac_results li", timeout=8000)
        page.locator("div.ac_results li").first.click()

        cursor = start
        while cursor <= end:
            page.fill("#srchSailDate", cursor.strftime("%Y-%m-%d"))
            try:
                page.select_option("#srchSelWeeks", value="8")
            except Exception:
                pass
            page.click("#btnRetrieve")
            page.wait_for_timeout(2500)
            texts = page.eval_on_selector_all(
                "#lsitContentArea2 .info-list-area > ul > li .result-info .list-result",
                "els => els.map(e => e.innerText)",
            )
            for t in texts:
                t = re.sub(r"\s+", " ", t).strip()
                m = re.search(
                    r"Origin\s+(?P<etd>\d{4}-\d{2}-\d{2})\s+.*?"
                    r"Destination\s+(?P<eta>\d{4}-\d{2}-\d{2})\s+.*?"
                    r"Main Vessel\s+(?P<vessel>.*?)\s+Route\s+(?P<route>\S+)\s+"
                    r"Operator\s+(?P<operator>\S+)", t)
                if m:
                    d = m.groupdict()
                    vm = re.match(r"^(.*)\((.*)\)$", d["vessel"].strip())
                    vessel, voyage = (vm.group(1).strip(), vm.group(2).strip()) if vm else (d["vessel"], "")
                    rows.append({
                        "carrier": "HMM", "vessel": vessel, "voyage": voyage,
                        "service": d["route"], "pol": "Shanghai", "etd": d["etd"],
                        "pod": route["short"], "eta": d["eta"], "transit_days": None,
                        "routing": "Direct", "cutoff": None,
                    })
            cursor += dt.timedelta(weeks=7)
        browser.close()
    return rows


def fetch_hmm_from_cache(route: dict) -> list[dict]:
    """Fall back to any HMM_*.xlsx already exported into this route's folder."""
    if pd is None:
        return []
    candidates = sorted(route["folder"].glob("HMM_*.xlsx")) + sorted(route["folder"].glob("HMM*schedule*.xlsx"))
    if not candidates:
        return []
    path = candidates[0]
    try:
        df = pd.read_excel(path, header=None)
    except Exception:
        return []
    header_row = None
    for i, val in enumerate(df.iloc[:, 0]):
        if str(val).strip().lower() == "week":
            header_row = i
            break
    if header_row is None:
        return []
    df = pd.read_excel(path, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        vessel_raw = str(r.get("Main Vessel", "")).strip()
        vm = re.match(r"^(.*)\((.*)\)$", vessel_raw)
        vessel, voyage = (vm.group(1).strip(), vm.group(2).strip()) if vm else (vessel_raw, "")
        etd = r.get("Origin Date (ETD)")
        eta = r.get("Destination Date (ETA)")
        if pd.isna(etd) or not vessel:
            continue
        rows.append({
            "carrier": "HMM", "vessel": vessel, "voyage": voyage,
            "service": r.get("Route"), "pol": "Shanghai",
            "etd": str(etd)[:10], "pod": route["short"], "eta": str(eta)[:10] if not pd.isna(eta) else "",
            "transit_days": r.get("Transit (days)"), "routing": r.get("Routing") or "Direct",
            "cutoff": r.get("Port Cut-off"),
        })
    print(f"    (HMM) using cached export {path.name}: {len(rows)} sailings")
    return rows


def _glob_ci(folder: Path, patterns: list[str]) -> list[Path]:
    """Case-insensitive glob (Windows-style file names like MaerskTH.csv / Ooclman.xlsx)."""
    found: dict[str, Path] = {}
    for p in folder.iterdir():
        if not p.is_file() or p.name.startswith("~$"):
            continue
        for pat in patterns:
            if re.fullmatch(pat.replace(".", r"\.").replace("*", ".*"), p.name, flags=re.I):
                found[p.name.lower()] = p
    return sorted(found.values())


def _to_iso(value, dayfirst=True) -> str:
    ts = pd.to_datetime(value, errors="coerce", dayfirst=dayfirst)
    return "" if pd.isna(ts) else ts.date().isoformat()


def _to_iso_dt(value, dayfirst=True) -> str:
    ts = pd.to_datetime(value, errors="coerce", dayfirst=dayfirst)
    return "" if pd.isna(ts) else ts.strftime("%Y-%m-%d %H:%M")


def _transit_days(text) -> int | None:
    if isinstance(text, (int, float)) and not pd.isna(text):
        return int(text)
    m = re.search(r"(\d+)\s*day", str(text), re.I)
    if m:
        return int(m.group(1))
    return int(text) if str(text).strip().isdigit() else None


def _read_csv_any(path: Path):
    for enc in ("utf-8-sig", "cp874", "cp1252"):
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            continue
    return None


def load_maersk_files(route: dict) -> list[dict]:
    """Maersk point-to-point results exported/copied from maersk.com (MaerskTH.csv, Maerskvn.csv,
    MaerskMan.csv ... or the older maersk_*.xlsx layout)."""
    if pd is None:
        return []
    rows: list[dict] = []
    files = _glob_ci(route["folder"], ["maersk*.csv", "maersk*.xlsx"])
    for path in files:
        df = _read_csv_any(path) if path.suffix.lower() == ".csv" else pd.read_excel(path)
        if df is None or df.empty:
            continue
        for _, r in df.iterrows():
            if "Voyage Number" in df.columns:  # new CSV layout
                vessel, voyage = str(r.get("Vessel", "")).strip(), str(r.get("Voyage Number", "")).strip()
                etd_col, eta_col = "Departure Date", "Arrival Date"
                cutoff = _to_iso_dt(r.get("Deadline CY"))
                pol_text, pod_text = str(r.get("Departure", "")), str(r.get("Arrival", ""))
            else:  # older maersk_schedule_export.py layout
                vv = str(r.get("Vessel / Voyage", r.get("Vessel", ""))).strip()
                vessel, voyage = _one_split_vessel_voyage(vv)
                etd_col, eta_col = "Departure Date", "Arrival Date"
                cutoff = ""
                pol_text, pod_text = str(r.get("Departure Port", "")), str(r.get("Arrival Port", ""))
            etd, eta = _to_iso(r.get(etd_col)), _to_iso(r.get(eta_col))
            if not vessel or not etd:
                continue
            rows.append({
                "carrier": "MAERSK", "vessel": vessel, "voyage": voyage, "service": None,
                "pol": "Shanghai", "etd": etd, "pod": route["short"], "eta": eta,
                "transit_days": _transit_days(r.get("Transit Time")), "routing": "Direct",
                "cutoff": cutoff, "note": pol_text.split(" - ", 1)[-1] + " -> " + pod_text.split(" - ", 1)[-1],
            })
    if rows:
        print(f"    (Maersk) {len(rows)} sailings from {', '.join(p.name for p in files)}")
    return rows


_OOCL_DATE = re.compile(r"(\d{1,2})\s+([A-Za-z]{3})")


def _oocl_date(text, year: int, not_before: dt.date | None = None) -> str:
    m = _OOCL_DATE.search(str(text))
    if not m:
        return ""
    try:
        d = dt.datetime.strptime(f"{m.group(1)} {m.group(2)} {year}", "%d %b %Y").date()
    except ValueError:
        return ""
    if not_before and d < not_before:
        d = d.replace(year=year + 1)
    return d.isoformat()


def load_oocl_files(route: dict) -> list[dict]:
    """OOCL 'Sailing Schedule' Excel export: every sailing is two rows (row 1 = CY cut-off / ETD /
    service / vessel, row 2 = ETD date / ETA date / SI cut-off). The year is only in the header line."""
    if pd is None:
        return []
    rows: list[dict] = []
    files = _glob_ci(route["folder"], ["oocl*.xlsx"])
    for path in files:
        raw = pd.read_excel(path, header=None)
        m = re.search(r"(\d{4})", str(raw.iat[0, 0]))
        year = int(m.group(1)) if m else TODAY.year
        hdr = next((i for i in range(len(raw)) if str(raw.iat[i, 0]).strip() == "Origin"), None)
        if hdr is None:
            continue
        df = pd.read_excel(path, header=hdr)
        recs = df.to_dict("records")
        i = 0
        while i < len(recs) - 1:
            a, b = recs[i], recs[i + 1]
            vv = str(a.get("Vessel Voyage", "")).strip()
            if str(a.get("Origin", "")).strip().lower() != "shanghai" or not vv or vv == "nan":
                i += 1
                continue
            etd_dt = _oocl_date(a.get("ETD at POL"), year)
            etd_time = re.search(r"(\d{2}:\d{2})", str(a.get("ETD at POL")))
            eta = _oocl_date(b.get("Destination"), year, not_before=dt.date.fromisoformat(etd_dt) if etd_dt else None)
            vessel, voyage = _one_split_vessel_voyage(vv)
            ts_port = a.get("Transshipment Port")
            is_ts = isinstance(ts_port, str) and ts_port.strip() != ""
            cut_txt = str(a.get("Cutoff", ""))
            cut_time = re.search(r"(\d{2}:\d{2})", cut_txt)
            cut_day = _oocl_date(cut_txt, year)
            cut = f"{cut_day} {cut_time.group(1)}" if cut_day and cut_time else ""
            rows.append({
                "carrier": "OOCL", "vessel": vessel, "voyage": voyage,
                "service": str(a.get("Service", "")).strip() or None, "pol": "Shanghai",
                "etd": etd_dt, "pod": route["short"], "eta": eta,
                "transit_days": _transit_days(a.get("Est. Transit Time")),
                "routing": "T/S" if is_ts else "Direct", "cutoff": cut,
                "note": f"T/S {ts_port}" if is_ts else "",
            })
            i += 2
    rows = [r for r in rows if r["etd"]]
    if rows:
        print(f"    (OOCL) {len(rows)} sailings from {', '.join(p.name for p in files)}")
    return rows


def load_hmm_html_files(route: dict) -> list[dict]:
    """HMM 'Retrieve -> Excel' download: the .xls is really an HTML table (Origin Point, Loading Port
    'SHANGHAI,CHINA ETD : 2026-09-22 09:00', Operator, Route, Vessel '[VTX]SM JAKARTA 2612W', ...)."""
    rows: list[dict] = []
    files = _glob_ci(route["folder"], ["hmm*.xls"])
    from bs4 import BeautifulSoup
    for path in files:
        raw = path.read_bytes()
        if raw[:4] == b"PK\x03\x04" or raw[:4] == b"\xd0\xcf\x11\xe0":
            continue  # a genuine Excel file, handled elsewhere
        soup = BeautifulSoup(raw.decode("utf-8-sig", errors="replace"), "html.parser")
        table = soup.find("table")
        if not table:
            continue
        heads = [th.get_text(" ", strip=True) for th in table.find_all("th")]
        for tr in table.find_all("tr"):
            tds = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
            if len(tds) != len(heads):
                continue
            rec = dict(zip(heads, tds))
            m_etd = re.search(r"ETD\s*:\s*(\d{4}-\d{2}-\d{2})", rec.get("Loading Port", ""))
            m_eta = re.search(r"ET[AB]\s*:\s*(\d{4}-\d{2}-\d{2})", rec.get("DischargingPort", ""))
            legs = re.findall(r"\[(\w+)\]\s*([^\[]+)", rec.get("Vessel", ""))
            if not m_etd or not legs:
                continue
            vessel, voyage = _one_split_vessel_voyage(legs[0][1].strip())
            ts_port = rec.get("Next Port(T/S)", "").strip()
            note = ""
            if len(legs) > 1:
                note = "Connecting vessel(s): " + " > ".join(f"{v.strip()}" for _, v in legs[1:])
                if ts_port:
                    note = f"T/S {ts_port} · " + note
            op, svc = rec.get("Operator", "").strip(), rec.get("Route", "").strip()
            days = rec.get("Total TransitTime(Days)", "").strip()
            rows.append({
                "carrier": "HMM", "vessel": vessel, "voyage": voyage,
                "service": f"{svc} ({op})" if op and svc else (svc or op or None),
                "pol": "Shanghai", "etd": m_etd.group(1), "pod": route["short"],
                "eta": m_eta.group(1) if m_eta else "",
                "transit_days": int(days) if days.isdigit() else None,
                "routing": "T/S" if ts_port or len(legs) > 1 else "Direct",
                "cutoff": _to_iso_dt(rec.get("Cargo Cut-off"), dayfirst=False), "note": note,
            })
    if rows:
        print(f"    (HMM) {len(rows)} sailings from {', '.join(p.name for p in files)}")
    return rows


def _read_table_after_marker(path: Path, marker: str):
    """Read an Excel sheet whose real header row starts with `marker` (title/notes rows above it)."""
    raw = pd.read_excel(path, header=None)
    hdr = next((i for i in range(len(raw)) if str(raw.iat[i, 0]).strip() == marker), None)
    if hdr is None:
        return None
    df = pd.read_excel(path, header=hdr)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _col(df, prefix: str):
    return next((c for c in df.columns if c.lower().startswith(prefix.lower())), None)


def _txt(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip().lower() in ("nan", "nat") else str(v).strip()


def load_hapag_files(route: dict) -> list[dict]:
    """Hapag-Lloyd schedule exported from hapag-lloyd.com (HapagLloyd_*.xlsx). Every sailing goes
    via a transshipment hub (column 'Via'), so all rows are T/S."""
    if pd is None:
        return []
    rows: list[dict] = []
    files = _glob_ci(route["folder"], ["hapag*.xlsx", "hapag*.xls"])
    for path in files:
        df = _read_table_after_marker(path, "No.")
        if df is None:
            continue
        c_etd, c_eta, c_days = _col(df, "ETD"), _col(df, "ETA"), _col(df, "Transit")
        c_via, c_cut = _col(df, "Via"), _col(df, "FCL Cut")
        for _, r in df.iterrows():
            etd, eta = _to_iso(r.get(c_etd), dayfirst=False), _to_iso(r.get(c_eta), dayfirst=False)
            vessel = _txt(r.get("Vessel"))
            if not etd or not vessel:
                continue
            via = _txt(r.get(c_via)) if c_via else ""
            routing_txt = _txt(r.get("Routing")).lower() if "Routing" in df.columns else ""
            is_ts = bool(via) or (bool(routing_txt) and not routing_txt.startswith("direct"))
            rows.append({
                "carrier": "HAPAG", "vessel": vessel, "voyage": _txt(r.get("Voyage")),
                "service": _txt(r.get("Service")) or None, "pol": "Shanghai", "etd": etd,
                "pod": route["short"], "eta": eta, "transit_days": _transit_days(r.get(c_days)),
                "routing": "T/S" if is_ts else "Direct", "cutoff": _to_iso_dt(r.get(c_cut), dayfirst=False),
                "note": f"T/S via {via}" if via else "",
            })
    if rows:
        print(f"    (Hapag-Lloyd) {len(rows)} sailings from {', '.join(p.name for p in files)}")
    return rows


def load_msc_files(route: dict) -> list[dict]:
    """MSC 'Search a Schedule' results copied into MSC_*.xlsx (ETD, ETA, Vessel, Voyage, Transit, Routing)."""
    if pd is None:
        return []
    rows: list[dict] = []
    files = _glob_ci(route["folder"], ["msc*.xlsx"])
    for path in files:
        df = _read_table_after_marker(path, "No.")
        if df is None:
            continue
        c_etd, c_eta, c_days = _col(df, "Departure"), _col(df, "Arrival"), _col(df, "Transit")
        for _, r in df.iterrows():
            etd, eta = _to_iso(r.get(c_etd), dayfirst=False), _to_iso(r.get(c_eta), dayfirst=False)
            vessel, voyage = _txt(r.get("Vessel")), _txt(r.get("Voyage"))
            if not etd or not vessel:
                continue
            note = ""
            if vessel.upper() == "TBN":
                vessel, note = f"TBN ({voyage})", "Vessel to be nominated"
            rows.append({
                "carrier": "MSC", "vessel": vessel, "voyage": voyage,
                "service": _txt(r.get("Service")) or None, "pol": "Shanghai", "etd": etd,
                "pod": route["short"], "eta": eta, "transit_days": _transit_days(r.get(c_days)),
                "routing": "Direct" if _txt(r.get("Routing")).lower().startswith("direct") else "T/S",
                "cutoff": "", "note": note,
            })
    if rows:
        print(f"    (MSC) {len(rows)} sailings from {', '.join(p.name for p in files)}")
    return rows


def load_one_files(route: dict) -> list[dict]:
    """ONE point-to-point export (ONE_Schedule_*.xlsx: ETD Shanghai, Vessel / Voyage, Service Lane,
    ETA Cai Mep, ETA Ho Chi Minh (Cat Lai), cut-offs). The file has no Direct/T-S column, so routing is
    inferred from transit time (> 11 days = transshipment) and flagged in the line note."""
    if pd is None:
        return []
    rows: list[dict] = []
    files = _glob_ci(route["folder"], ["one_schedule*.xlsx"])
    for path in files:
        df = _read_table_after_marker(path, "ETD Shanghai")
        if df is None:
            continue
        eta_cols = [c for c in df.columns if c.startswith("ETA ")]
        day_cols = [c for c in df.columns if c.lower().startswith("transit ")]
        # prefer the column for the actual destination (e.g. Cat Lai) over the nearby port (Cai Mep)
        eta_cols.sort(key=lambda c: 0 if "Cat Lai" in c or route["short"].lower() in c.lower() else 1)
        for _, r in df.iterrows():
            etd = _to_iso(r.get("ETD Shanghai"), dayfirst=False)
            vv = _txt(r.get("Vessel / Voyage"))
            if not etd or not vv:
                continue
            eta, days, used = "", None, ""
            for ec in eta_cols:
                iso = _to_iso(r.get(ec), dayfirst=False)
                if iso:
                    eta, used = iso, ec
                    days = (dt.date.fromisoformat(eta) - dt.date.fromisoformat(etd)).days
                    break
            if not eta:
                continue
            vessel, voyage = _one_split_vessel_voyage(vv)
            is_ts = days > 11
            note = "Routing inferred from transit time" + (f" · ETA at {used[4:]}" if "Cat Lai" not in used else "")
            rows.append({
                "carrier": "ONE", "vessel": vessel, "voyage": voyage,
                "service": _txt(r.get("Service Lane")) or None, "pol": "Shanghai", "etd": etd,
                "pod": route["short"], "eta": eta, "transit_days": days,
                "routing": "T/S" if is_ts else "Direct",
                "cutoff": _to_iso_dt(r.get("Port Cut-off"), dayfirst=False), "note": note,
            })
    if rows:
        print(f"    (ONE) {len(rows)} sailings from {', '.join(p.name for p in files)}")
    return rows


def fetch_msc(route: dict, start: dt.date, end: dt.date) -> list[dict]:
    """MSC public search API (see msc_schedule_*.py). msc.com currently answers 403 / Access Denied
    to automated clients from most networks, in which case this raises and the carrier is marked blocked."""
    base = "https://www.msc.com"
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*", "Content-Type": "application/json",
        "X-Requested-With": "XMLHttpRequest", "Origin": base, "Referer": f"{base}/en/search-a-schedule",
    })
    s.get(f"{base}/en/search-a-schedule", timeout=20)
    payload = {"FromDate": max(start, TODAY).isoformat(), "fromPortId": 444, "toPortId": route["msc_port_id"],
               "language": "en", "dataSourceId": "{E9CCBD25-6FBA-4C5C-85F6-FC4F9E5A931F}"}
    r = s.post(f"{base}/api/feature/tools/SearchSailingRoutes", json=payload, timeout=30)
    r.raise_for_status()
    rows = []
    for group in (r.json().get("Data") or []):
        for rt in group.get("Routes", []):
            rows.append({
                "carrier": "MSC", "vessel": rt.get("VesselName"), "voyage": rt.get("DepartureVoyageNo"),
                "service": group.get("LoadingService"), "pol": "Shanghai", "pod": route["short"],
                "etd": (rt.get("EstimatedDepartureDate") or "")[:10], "eta": (rt.get("EstimatedArrivalDate") or "")[:10],
                "transit_days": rt.get("TotalTransitTime"),
                "routing": "Direct" if str(group.get("RoutingType", "")).lower().startswith("direct") else "T/S",
                "cutoff": (rt.get("CutOffs") or {}).get("ContainerYardCutOffDate"),
            })
    return rows


# ---------------------------------------------------------------------------
# Normalisation + merge (collapse the same physical vessel call across lines)
# ---------------------------------------------------------------------------
def normalize_vessel(name: str) -> str:
    name = (name or "").upper().strip()
    name = re.sub(r"[^A-Z0-9 ]", "", name)
    name = re.sub(r"\s+", " ", name)
    return name


MERGE_WINDOW_DAYS = 2  # carriers publish ETD in different time zones / berth windows


def merge_route_rows(route_code: str, route_name: str, rows: list[dict]) -> list[dict]:
    """One calendar entry per physical vessel call: same vessel name and ETD within +/-2 days.
    Every carrier selling that call is listed in `lines` (one line per carrier+voyage)."""
    def _iso(v):
        m = re.match(r"\s*(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(v or ""))
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else ""

    rows = [{**r, "etd": _iso(r.get("etd")), "eta": _iso(r.get("eta"))} for r in rows]
    by_vessel: dict[str, list[dict]] = {}
    for r in sorted((x for x in rows if x.get("vessel") and x.get("etd")), key=lambda x: x["etd"]):
        by_vessel.setdefault(normalize_vessel(r["vessel"]), []).append(r)

    merged: list[dict] = []
    for vessel_key, vrows in by_vessel.items():
        clusters: list[list[dict]] = []
        for r in vrows:
            d = dt.date.fromisoformat(r["etd"])
            if clusters and (d - dt.date.fromisoformat(clusters[-1][0]["etd"])).days <= MERGE_WINDOW_DAYS:
                clusters[-1].append(r)
            else:
                clusters.append([r])
        for cl in clusters:
            etds = [c["etd"] for c in cl]
            etd = max(set(etds), key=lambda e: (etds.count(e), -dt.date.fromisoformat(e).toordinal()))
            lines: list[dict] = []
            for c in cl:
                line = {
                    "carrier": c["carrier"], "voyage": c.get("voyage") or "", "service": c.get("service") or "",
                    "cutoff": c.get("cutoff") or "", "etd": c["etd"], "eta": c.get("eta") or "",
                    "routing": c.get("routing") or "Direct", "note": c.get("note") or "",
                    "transit_days": _transit_days(c.get("transit_days")) if c.get("transit_days") not in (None, "") else None,
                }
                same = next((i for i, l in enumerate(lines)
                             if l["carrier"] == line["carrier"] and l["voyage"].strip() == line["voyage"].strip()), None)
                if same is None:
                    lines.append(line)
                elif lines[same]["routing"] != "Direct" and line["routing"] == "Direct":
                    lines[same] = line  # keep the direct variant of the same carrier voyage
            direct = [c for c in cl if c.get("routing") == "Direct"]
            pick = (direct or cl)[0]
            eta = pick.get("eta") or next((c["eta"] for c in cl if c.get("eta")), "")
            days = pick.get("transit_days")
            try:
                days = int(float(days))
            except (TypeError, ValueError):
                days = None
            if days is None and eta:
                days = (dt.date.fromisoformat(eta) - dt.date.fromisoformat(etd)).days
            merged.append({
                "route_code": route_code, "route_name": route_name,
                "vessel": cl[0]["vessel"].strip(), "etd": etd, "eta": eta, "transit_days": days,
                "routing": "Direct" if direct else "T/S", "pol": "Shanghai", "pod": route_name,
                "lines": lines,
            })
    merged.sort(key=lambda g: (g["etd"], g["vessel"]))
    return merged


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Build the merged Shanghai sailing-schedule calendar/dashboard.")
    ap.add_argument("--start", default=DEFAULT_START.isoformat())
    ap.add_argument("--end", default=DEFAULT_END.isoformat())
    ap.add_argument("--use-cache", action="store_true",
                     help="Reuse any *_raw.json already saved per route instead of re-fetching COSCO/Yang Ming/ONE live.")
    ap.add_argument("--skip-one", action="store_true", help="Skip the (slower) ONE Playwright scrape entirely.")
    ap.add_argument("--skip-hmm-live", action="store_true", help="Don't attempt a live HMM fetch, go straight to cache.")
    args = ap.parse_args()

    start = dt.date.fromisoformat(args.start)
    end = dt.date.fromisoformat(args.end)

    DASHBOARD_DIR.mkdir(exist_ok=True)
    ASSETS_DIR.mkdir(exist_ok=True)

    all_merged = []
    carrier_status = {}

    for route in ROUTES:
        print(f"\n=== {route['code']} : Shanghai -> {route['name']} ===")
        live_path = route["folder"] / "route_live_rows.json"
        live_rows: list[dict] = []
        status: dict[str, str] = {}

        if args.use_cache and live_path.exists():
            cached = json.loads(live_path.read_text(encoding="utf-8"))
            live_rows, status = cached["rows"], cached["status"]
            print(f"  (live carriers COSCO/Yang Ming/ONE from cache: {len(live_rows)} rows)")
        else:
            for code, label, fn in (("COSCO", "COSCO", fetch_cosco), ("YML", "Yang Ming", fetch_yangming),
                                    ("ONE", "ONE", fetch_one)):
                if code == "ONE" and args.skip_one:
                    status[code] = "unavailable"
                    continue
                try:
                    rows = fn(route, start, end)
                    print(f"  {label + ':':<11}{len(rows)} sailings (live)")
                    live_rows.extend(rows)
                    status[code] = "live" if rows else "unavailable"
                except Exception as e:
                    print(f"  {label + ':':<11}FAILED ({type(e).__name__}: {str(e)[:100]})")
                    status[code] = "unavailable"
            live_path.write_text(json.dumps({"rows": live_rows, "status": status}, ensure_ascii=False, indent=2),
                                 encoding="utf-8")

        route_rows = list(live_rows)

        # Files copied/exported from the carriers' own sites (always re-read, so new uploads are picked up)
        file_sources = {
            "HMM": lambda: load_hmm_html_files(route) or fetch_hmm_from_cache(route),
            "MAERSK": lambda: load_maersk_files(route),
            "OOCL": lambda: load_oocl_files(route),
            "MSC": lambda: load_msc_files(route),
            "HAPAG": lambda: load_hapag_files(route),
            "ONE": lambda: load_one_files(route),
        }
        for code, loader in file_sources.items():
            rows = loader()
            route_rows.extend(rows)
            if rows:
                status[code] = "file"
            else:
                status.setdefault(code, "unavailable")

        if status.get("MSC") != "file":
            try:
                msc_rows = fetch_msc(route, start, end)
                route_rows.extend(msc_rows)
                status["MSC"] = "live" if msc_rows else "unavailable"
            except Exception as e:
                print(f"  MSC:       blocked/unavailable ({type(e).__name__})")
                status["MSC"] = "blocked"

        for code, st in status.items():
            carrier_status.setdefault(code, {})[route["code"]] = st

        merged = merge_route_rows(route["code"], route["short"], route_rows)
        print(f"  -> {len(route_rows)} raw sailings merged into {len(merged)} calendar entries")
        all_merged.extend(merged)

    # ---- write dataset ----
    dataset = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "period": {"start": start.isoformat(), "end": end.isoformat()},
        "routes": [{"code": r["code"], "name": r["name"], "short": r["short"]} for r in ROUTES],
        "carriers": CARRIERS,
        "carrier_status": carrier_status,
        "sailings": all_merged,
    }
    (DASHBOARD_DIR / "data.json").write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nSaved dataset: {DASHBOARD_DIR / 'data.json'} ({len(all_merged)} total calendar entries)")

    # ---- copy logo assets ----
    logo_srcs = {
        "cosco.png": BASE_DIR / "cosco.png",
        "maersk.png": BASE_DIR / "marsk.png",
        "hmm.webp": BASE_DIR / "HMM_Logo_Basic_Form.svg.webp",
        "one.webp": BASE_DIR / "Ocean_Network_Express_logo.svg.webp",
        "yangming.png": BASE_DIR / "yangming.png",
        "oocl.png": BASE_DIR / "oocl-logo.png",
        "msc.jpg": BASE_DIR / "msc.jpg",
        "hapag.png": BASE_DIR / "hapag.png",
    }
    for dest_name, src in logo_srcs.items():
        if src.exists():
            shutil.copy(src, ASSETS_DIR / dest_name)

    print(f"Dashboard folder ready: {DASHBOARD_DIR}")
    render_html()
    print("Open dashboard/index.html in a browser to view it.")


def render_html():
    """Inject dashboard/data.json into the HTML template to produce index.html."""
    template_path = DASHBOARD_DIR / "index.template.html"
    data_path = DASHBOARD_DIR / "data.json"
    out_path = DASHBOARD_DIR / "index.html"

    template = template_path.read_text(encoding="utf-8")
    data_json = data_path.read_text(encoding="utf-8")
    html = template.replace("__DATA_JSON__", data_json.replace("</", "<\\/"))
    out_path.write_text(html, encoding="utf-8")
    print(f"Rendered: {out_path}")


if __name__ == "__main__":
    if "--render-only" in sys.argv:
        render_html()
    else:
        main()
