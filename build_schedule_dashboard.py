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

Carrier coverage in this build
-------------------------------
  COSCO     - live: elines.coscoshipping.com public JSON API (fast, reliable)
  Yang Ming - live: yangming.com public JSON API (fast, reliable)
  ONE       - live: one-line.com point-to-point page, scraped with Playwright
              (the site's own "Download" button flow is fragile because of a
              cookie-consent overlay, so this reads the rendered result cards
              directly instead - see fetch_one())
  HMM       - best-effort live (Playwright); hmm21.com was unreachable from
              this machine at the time this was written, so if the live
              fetch fails this falls back to any previously-exported
              HMM_*.xlsx already sitting in that route's folder
  Maersk    - best-effort live (Playwright); maersk.com's schedule page is a
              client-side app and the deep-link URL scheme the original
              maersk_schedule_export.py scripts relied on no longer returns
              results (it now renders the generic search landing page), so
              this carrier is skipped with a warning unless a cached
              maersk_*.xlsx is already present

Usage
-----
    python build_schedule_dashboard.py                 # fetch + rebuild everything
    python build_schedule_dashboard.py --use-cache      # reuse any *_raw.json / *.xlsx already on disk, skip live fetches
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
        "hmm_dest": "MANZANILLO",
        "hmm_dest_code": "MXZLO",
    },
]

CARRIERS = {
    "COSCO": {"label": "COSCO Shipping", "color": "#1a3d8f", "logo": "cosco.png"},
    "HMM": {"label": "HMM", "color": "#0a2a66", "logo": "hmm.webp"},
    "ONE": {"label": "Ocean Network Express (ONE)", "color": "#e2007a", "logo": "one.webp"},
    "YML": {"label": "Yang Ming", "color": "#0f7a3d", "logo": None},
    "MAERSK": {"label": "Maersk", "color": "#42b0d5", "logo": "maersk.png"},
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


def fetch_maersk_from_cache(route: dict) -> list[dict]:
    if pd is None:
        return []
    candidates = sorted(route["folder"].glob("maersk_*.xlsx"))
    if not candidates:
        return []
    path = candidates[0]
    try:
        df = pd.read_excel(path)
    except Exception:
        return []
    rows = []
    vv_col = "Vessel / Voyage" if "Vessel / Voyage" in df.columns else "Vessel"
    for _, r in df.iterrows():
        vv = str(r.get(vv_col, "")).strip()
        vessel, voyage = _one_split_vessel_voyage(vv)
        etd = str(r.get("Departure Date", "")).strip()
        eta = str(r.get("Arrival Date", "")).strip()
        if not vessel or not etd:
            continue
        rows.append({
            "carrier": "MAERSK", "vessel": vessel, "voyage": voyage,
            "service": None, "pol": r.get("Departure Port") or "Shanghai",
            "etd": etd, "pod": r.get("Arrival Port") or route["short"], "eta": eta,
            "transit_days": None, "routing": "Direct", "cutoff": None,
        })
    print(f"    (Maersk) using cached export {path.name}: {len(rows)} sailings")
    return rows


# ---------------------------------------------------------------------------
# Normalisation + merge (collapse the same physical vessel/date across lines)
# ---------------------------------------------------------------------------
def normalize_vessel(name: str) -> str:
    name = (name or "").upper().strip()
    name = re.sub(r"[^A-Z0-9 ]", "", name)
    name = re.sub(r"\s+", " ", name)
    return name


def merge_route_rows(route_code: str, route_name: str, rows: list[dict]) -> list[dict]:
    groups: dict[tuple, dict] = {}
    order: list[tuple] = []
    for r in rows:
        if not r.get("vessel") or not r.get("etd"):
            continue
        key = (normalize_vessel(r["vessel"]), r["etd"])
        if key not in groups:
            groups[key] = {
                "route_code": route_code,
                "route_name": route_name,
                "vessel": r["vessel"].strip(),
                "etd": r["etd"],
                "eta": r.get("eta") or "",
                "transit_days": r.get("transit_days"),
                "routing": r.get("routing") or "Direct",
                "pol": r.get("pol") or "Shanghai",
                "pod": r.get("pod") or route_name,
                "lines": [],
            }
            order.append(key)
        g = groups[key]
        g["lines"].append({
            "carrier": r["carrier"],
            "voyage": r.get("voyage") or "",
            "service": r.get("service") or "",
            "cutoff": r.get("cutoff") or "",
        })
        if not g["eta"] and r.get("eta"):
            g["eta"] = r["eta"]
        if g.get("transit_days") in (None, "") and r.get("transit_days") not in (None, ""):
            g["transit_days"] = r["transit_days"]
        if r.get("routing") == "Direct":
            g["routing"] = "Direct"

    merged = [groups[k] for k in order]
    merged.sort(key=lambda g: g["etd"])
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
        route_rows = []
        raw_cache_path = route["folder"] / "route_raw_rows.json"

        cached_raw = None
        if args.use_cache and raw_cache_path.exists():
            cached_raw = json.loads(raw_cache_path.read_text(encoding="utf-8"))

        if cached_raw is not None:
            route_rows = cached_raw
            print(f"  (using cached route_raw_rows.json: {len(route_rows)} rows)")
        else:
            # COSCO
            try:
                rows = fetch_cosco(route, start, end)
                print(f"  COSCO:     {len(rows)} sailings")
                route_rows.extend(rows)
                carrier_status.setdefault("COSCO", {})[route["code"]] = "live"
            except Exception as e:
                print(f"  COSCO:     FAILED ({e})")
                carrier_status.setdefault("COSCO", {})[route["code"]] = "unavailable"

            # Yang Ming
            try:
                rows = fetch_yangming(route, start, end)
                print(f"  Yang Ming: {len(rows)} sailings")
                route_rows.extend(rows)
                carrier_status.setdefault("YML", {})[route["code"]] = "live"
            except Exception as e:
                print(f"  Yang Ming: FAILED ({e})")
                carrier_status.setdefault("YML", {})[route["code"]] = "unavailable"

            # ONE
            if not args.skip_one:
                try:
                    rows = fetch_one(route, start, end)
                    print(f"  ONE:       {len(rows)} sailings")
                    route_rows.extend(rows)
                    carrier_status.setdefault("ONE", {})[route["code"]] = "live"
                except Exception as e:
                    print(f"  ONE:       FAILED ({e})")
                    carrier_status.setdefault("ONE", {})[route["code"]] = "unavailable"
            else:
                carrier_status.setdefault("ONE", {})[route["code"]] = "skipped"

            # HMM: live attempt, else cache
            hmm_rows = []
            if not args.skip_hmm_live:
                try:
                    hmm_rows = fetch_hmm_live(route, start, end)
                    print(f"  HMM:       {len(hmm_rows)} sailings (live)")
                    carrier_status.setdefault("HMM", {})[route["code"]] = "live"
                except Exception as e:
                    print(f"  HMM:       live fetch failed ({type(e).__name__}: {str(e)[:120]})")
            if not hmm_rows:
                hmm_rows = fetch_hmm_from_cache(route)
                carrier_status.setdefault("HMM", {})[route["code"]] = "cached" if hmm_rows else "unavailable"
            route_rows.extend(hmm_rows)

            # Maersk: cache only (live deep-link automation is currently broken on maersk.com)
            maersk_rows = fetch_maersk_from_cache(route)
            carrier_status.setdefault("MAERSK", {})[route["code"]] = "cached" if maersk_rows else "unavailable"
            route_rows.extend(maersk_rows)

            raw_cache_path.write_text(json.dumps(route_rows, ensure_ascii=False, indent=2), encoding="utf-8")

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
    html = template.replace("__DATA_JSON__", data_json)
    out_path.write_text(html, encoding="utf-8")
    print(f"Rendered: {out_path}")


if __name__ == "__main__":
    if "--render-only" in sys.argv:
        render_html()
    else:
        main()
