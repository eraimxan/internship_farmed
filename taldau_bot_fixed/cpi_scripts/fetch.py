"""
cpi_scripts/fetch.py — Fetch Consumer Price Index data from taldau.stat.gov.kz

Source:  https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/703076?keyword=инфляция
Index:   703076 — Индексы потребительских цен (CPI)
Periods: "отчетный период к соответствующему периоду прошлого года" (YoY)
         "отчетный период к предыдущему году" (vs previous full year)

Usage:
    python fetch.py                  # uses cookie from ../data/cookie.txt
    python fetch.py --cookie "..."   # provide cookie manually
    python fetch.py --out mydata.csv # custom output path

API parameters (discovered once, hardcoded for speed + accuracy):
    period_id    = 4  (Месяц — monthly; both comparison types available only here)
    dic_ids      = 67,848,2753  (Regions, Comparison, Goods)
    goods_root   = 4772381  ("Товары и услуги")
    goods_class_root = 7358  (hidden wrapper, drilled past)
    yoy_term     = 2695732  (к соотв. периоду прошлого года)
    prev_period  = 2695730  (к предыдущему периоду)
    21 regions   = oblast-level IDs (CPI not available at district level)

Outputs:
    output/cpi_data.csv     — full data, all periods, both comparison types
    output/cpi_params.json  — discovered period names and metadata
"""

import sys
import re
import time
import json
import logging
import argparse
import threading
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

from tqdm import tqdm

import requests
import pandas as pd

# Allow running as standalone
sys.path.insert(0, str(Path(__file__).resolve().parent))
from cpi_config import (
    INDEX_ID, BASE_URL, API_BASE,
    COOKIE_FILE, COOKIE_STRING,
    REQUEST_DELAY, RETRY_COUNT, REQUEST_TIMEOUT, RETRY_BACKOFF_BASE, MAX_DEPTH,
    TARGET_PERIODS, CPI_REGIONS, DATA_CSV, CHECKPOINT_CSV, KNOWN_PARAMS,
    FETCH_WORKERS,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger(__name__)

# Signals all worker threads to stop (set on Ctrl+C)
_stop_event = threading.Event()

# Live request counter shared across worker threads (for progress bar postfix)
_req_lock        = threading.Lock()
_req_count       = 0
_req_bar: "tqdm | None" = None
_active_regions: dict[str, str] = {}   # region_name -> current comp_key

# Diagnostic: count tree-walk calls per depth (reveals tree explosion)
_depth_counts: dict[int, int] = {}
_depth_lock = threading.Lock()


def _bump_request(success: bool = True) -> None:
    """Called from each worker after every HTTP attempt — keeps the bar alive."""
    global _req_count
    with _req_lock:
        if success:
            _req_count += 1
        if _req_bar is not None:
            active = ", ".join(f"{n[:6]}…[{c}]" for n, c in list(_active_regions.items())[:2])
            _req_bar.set_postfix_str(f"reqs={_req_count} | {active}", refresh=True)


def _set_active(region_name: str, comp_key: str | None) -> None:
    with _req_lock:
        if comp_key is None:
            _active_regions.pop(region_name, None)
        else:
            _active_regions[region_name] = comp_key
        if _req_bar is not None:
            active = ", ".join(f"{n[:6]}…[{c}]" for n, c in list(_active_regions.items())[:2])
            _req_bar.set_postfix_str(f"reqs={_req_count} | {active}", refresh=True)


# ─────────────────────────────────────────────
# SESSION
# ─────────────────────────────────────────────

HEADERS = {
    "User-Agent":       "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 Chrome/123.0.0.0 Safari/537.36",
    "Accept":           "application/json, text/javascript, */*; q=0.01",
    "Accept-Language":  "ru-RU,ru;q=0.9,en;q=0.8",
    "X-Requested-With": "XMLHttpRequest",
    "Referer":          f"https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/{INDEX_ID}",
}


def _load_cookie() -> str:
    if COOKIE_STRING:
        return COOKIE_STRING
    if COOKIE_FILE.exists():
        c = COOKIE_FILE.read_text(encoding="utf-8").strip()
        if c:
            return c
    return ""


def make_session(cookie: str = "") -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    if cookie:
        s.headers["Cookie"] = cookie
    return s


def _get_periods_url(params: dict) -> str:
    measure_id = params.get("measure_id", "7")
    region_id  = CPI_REGIONS[0]["id"]
    comp_yoy   = params["comparison_yoy"]
    goods_root = params["goods_root"]
    period_id  = params["period_id"]
    dic_ids    = params["dic_ids"]
    all_terms  = f"{region_id},{comp_yoy},{goods_root}"
    return (
        f"{API_BASE}/GetIndexPeriods"
        f"?p_measure_id={measure_id}&p_index_id={INDEX_ID}&p_period_id={period_id}"
        f"&p_terms={all_terms}&p_term_id={goods_root}&p_dicIds={dic_ids}"
    )


def check_connection(session: requests.Session, params: dict) -> bool:
    """
    Validate cookie and API connectivity before starting the full fetch.
    Retries on 503 (server temporarily unavailable) up to RETRY_COUNT times.
    """
    tqdm.write("Checking connection to taldau.stat.gov.kz …")
    url = _get_periods_url(params)

    for attempt in range(RETRY_COUNT):
        try:
            r = session.get(url, timeout=REQUEST_TIMEOUT)
            if r.status_code in (301, 302, 401, 403):
                tqdm.write(f"  [FAIL] HTTP {r.status_code} — cookie is invalid or expired")
                return False
            if r.status_code == 503:
                wait = RETRY_BACKOFF_BASE * (attempt + 1)
                tqdm.write(f"  [WAIT] HTTP 503 — server busy, retrying in {wait}s "
                           f"(attempt {attempt + 1}/{RETRY_COUNT})")
                time.sleep(wait)
                continue
            if r.status_code != 200:
                tqdm.write(f"  [FAIL] HTTP {r.status_code} — unexpected response")
                return False
            data = r.json()
            dates = data.get("dateList", [])
            if not dates:
                tqdm.write("  [FAIL] Connected but got empty data — cookie may be expired")
                return False
            tqdm.write(f"  [ OK ] Connected — {len(dates)} periods available "
                       f"(latest: {data.get('periodNameList', ['?'])[-1]})")
            return True
        except Exception as exc:
            wait = RETRY_BACKOFF_BASE * (attempt + 1)
            tqdm.write(f"  [WAIT] Connection error: {exc} — retrying in {wait}s")
            time.sleep(wait)

    tqdm.write(f"  [FAIL] Server unavailable after {RETRY_COUNT} attempts — try again later")
    return False


# ─────────────────────────────────────────────
# LOW-LEVEL API CALL
# ─────────────────────────────────────────────

def _tree_post(
    session: requests.Session,
    period_id: str,
    dic_ids: str,
    p_terms: str,
    p_term_id: str,
    parent_id: str,
    idx: str = "0",
    measure_id: str = "7",
) -> list[dict]:
    """Single POST to GetIndexTreeData with retry logic."""
    data = {
        "p_parent_id":  parent_id,
        "p_index_id":   INDEX_ID,
        "p_keyword":    "",
        "p_period_id":  period_id,
        "p_measure_id": measure_id,
        "p_term_id":    p_term_id,
        "p_terms":      p_terms,
        "p_dicIds":     dic_ids,
        "idx":          idx,
        "filter":       '[{"property":null,"value":null}]',
        "id":           parent_id if parent_id else p_term_id,
    }
    for attempt in range(RETRY_COUNT):
        if _stop_event.is_set():
            return []
        try:
            r = session.post(BASE_URL, data=data, timeout=REQUEST_TIMEOUT)
            time.sleep(REQUEST_DELAY)
            if r.status_code != 200:
                wait = RETRY_BACKOFF_BASE * (attempt + 1)
                tqdm.write(f"  [WARN] HTTP {r.status_code} on attempt {attempt + 1} — retrying in {wait}s")
                time.sleep(wait)
                continue
            text = r.text.strip()
            if not text or text in ("[]", "null"):
                if attempt < RETRY_COUNT - 1:
                    time.sleep(RETRY_BACKOFF_BASE * (attempt + 1))
                    continue
                return []
            _bump_request(success=True)
            return r.json()
        except Exception as exc:
            wait = RETRY_BACKOFF_BASE * (attempt + 1)
            tqdm.write(f"  [WARN] Connection error attempt {attempt + 1}/{RETRY_COUNT}: {exc} — retrying in {wait}s")
            time.sleep(wait)
    tqdm.write(f"  [ERROR] All {RETRY_COUNT} attempts failed for parent_id={parent_id!r}")
    return []


# ─────────────────────────────────────────────
# PERIOD NAME DISCOVERY
# ─────────────────────────────────────────────

def get_period_names(session: requests.Session, params: dict) -> dict:
    """
    Call GetIndexPeriods to get the human-readable names for all period keys.
    Returns {period_key: "Jan 2024", ...}
    """
    period_id    = params["period_id"]
    dic_ids      = params["dic_ids"]
    goods_root   = params["goods_root"]
    comp_yoy     = params["comparison_yoy"]
    region_id    = CPI_REGIONS[0]["id"]   # use first region for discovery

    all_terms = f"{region_id},{comp_yoy},{goods_root}"

    url = _get_periods_url(params)

    for attempt in range(RETRY_COUNT):
        try:
            r = session.get(url, timeout=REQUEST_TIMEOUT)
            if r.status_code == 503:
                wait = RETRY_BACKOFF_BASE * (attempt + 1)
                log.warning("GetIndexPeriods: HTTP 503 — retrying in %ds (attempt %d/%d)",
                            wait, attempt + 1, RETRY_COUNT)
                time.sleep(wait)
                continue
            r.raise_for_status()
            data = r.json()
            date_list = data.get("dateList",       [])
            name_list = data.get("periodNameList", [])
            result = {f"y{k}": n for k, n in zip(date_list, name_list)}
            log.info("Period names: %d entries", len(result))
            return result, date_list, name_list
        except Exception as exc:
            wait = RETRY_BACKOFF_BASE * (attempt + 1)
            log.warning("GetIndexPeriods failed (attempt %d/%d): %s — retrying in %ds",
                        attempt + 1, RETRY_COUNT, exc, wait)
            time.sleep(wait)

    log.warning("GetIndexPeriods gave up after %d attempts — period labels will be missing", RETRY_COUNT)
    return {}, [], []


# ─────────────────────────────────────────────
# RECURSIVE GOODS TREE FETCHER
# ─────────────────────────────────────────────

def _clean_val(v) -> float | str | None:
    s = str(v).strip()
    if s.lower() in ("x", "х"):
        return "x"
    if s.lower() in ("", "nan", "none", "null"):
        return None
    try:
        return float(s.replace(" ", "").replace(",", "."))
    except Exception:
        return None


def fetch_tree(
    session: requests.Session,
    params: dict,
    region_id: str,
    comparison_term_id: str,
    parent_id: str,
    depth: int = 0,
    parent_text: str = "",
    visited: set[str] | None = None,
) -> list[dict]:
    """
    Recursively fetch the CPI goods tree for one (region, comparison_type) pair.

    The goods tree has a hidden wrapper root (id=7358, hidden=true).
    depth=0 starts at p_parent_id="" which returns the hidden wrapper.
    We recurse into it, so depth=1 returns "Товары и услуги" (4772381).
    We filter out the hidden wrapper row from the result.

    *visited* is a per-walk set of parent_ids already fetched; prevents cycles
    and cross-references from causing duplicate work. Initialized once per
    (region, comparison_type) pair and passed through recursion.
    """
    if depth > MAX_DEPTH:
        return []

    if visited is None:
        visited = set()
    if parent_id and parent_id in visited:
        return []
    if parent_id:
        visited.add(parent_id)

    # Diagnostic: track depth distribution + periodic summary
    with _depth_lock:
        _depth_counts[depth] = _depth_counts.get(depth, 0) + 1
        total = sum(_depth_counts.values())
        if total % 500 == 0:
            summary = " ".join(f"d{d}={_depth_counts[d]}" for d in sorted(_depth_counts))
            tqdm.write(f"  [DEPTH] {total} calls total: {summary}")

    p_terms = f"{region_id},{comparison_term_id},{params['goods_root']}"

    rows = _tree_post(
        session,
        period_id=params["period_id"],
        dic_ids=params["dic_ids"],
        p_terms=p_terms,
        p_term_id=params["goods_root"],
        parent_id=parent_id,
        idx="0",
        measure_id=params.get("measure_id", "7"),
    )

    result: list[dict] = []
    for row in rows:
        # Skip the hidden wrapper root
        if str(row.get("hidden", "false")).lower() == "true":
            # But still recurse into it to get the actual data
            child_rows = fetch_tree(
                session, params, region_id, comparison_term_id,
                parent_id=str(row["id"]),
                depth=depth,                  # don't increment depth for hidden node
                parent_text=parent_text,
                visited=visited,
            )
            result.extend(child_rows)
            continue

        period_keys = {
            k: _clean_val(v)
            for k, v in row.items()
            if re.match(r"^y\d{6}$", k)      # yMMYYYY format (e.g. y032025, y122024)
        }
        entry = {
            "id":          row.get("id"),
            "text":        str(row.get("text", "")).strip(),
            "leaf":        row.get("leaf", "true"),
            "depth":       depth,
            "parent_text": parent_text,
            **period_keys,
        }
        result.append(entry)

        if str(row.get("leaf", "true")).lower() == "false":
            result.extend(
                fetch_tree(
                    session, params, region_id, comparison_term_id,
                    parent_id=str(row["id"]),
                    depth=depth + 1,
                    parent_text=entry["text"],
                    visited=visited,
                )
            )
    return result


# ─────────────────────────────────────────────
# MAIN FETCH LOOP
# ─────────────────────────────────────────────

def _load_checkpoint(checkpoint_path: Path) -> tuple[pd.DataFrame, set[str]]:
    """
    Load existing checkpoint CSV.
    Returns (df_done, done_regions) where done_regions is the set of oblast names
    that were fully fetched (both comparison types present).
    """
    if not checkpoint_path.exists():
        return pd.DataFrame(), set()
    try:
        df = pd.read_csv(checkpoint_path, encoding="utf-8-sig", low_memory=False)
        if df.empty or "oblast" not in df.columns or "comparison_type" not in df.columns:
            return pd.DataFrame(), set()
        # A region is "done" only if BOTH comparison types are present for it
        comp_counts = df.groupby("oblast")["comparison_type"].nunique()
        done = set(comp_counts[comp_counts >= 2].index.tolist())
        if done:
            log.info("Checkpoint: %d region(s) already complete — skipping: %s",
                     len(done), ", ".join(sorted(done)))
        return df, done
    except Exception as exc:
        log.warning("Could not read checkpoint (%s) — starting fresh", exc)
        return pd.DataFrame(), set()


_checkpoint_lock = threading.Lock()


def _append_checkpoint(checkpoint_path: Path, rows: list[dict]) -> None:
    """Thread-safe append of one completed region's rows to the checkpoint CSV."""
    if not rows:
        return
    new_df = pd.DataFrame(rows)
    with _checkpoint_lock:
        write_header = not checkpoint_path.exists()
        new_df.to_csv(
            checkpoint_path,
            mode="a",
            index=False,
            encoding="utf-8-sig",
            header=write_header,
        )


def _fetch_region(
    cookie: str,
    params: dict,
    region: dict,
    index: int,
    total: int,
) -> tuple[str, list[dict], bool]:
    """
    Fetch both comparison types for a single region.
    Each worker gets its own session to avoid cross-thread connection sharing.
    Returns (region_name, rows, all_ok).
    """
    session = make_session(cookie)
    region_id   = region["id"]
    region_name = region["name"]
    comp_map = {
        "yoy":       params["comparison_yoy"],
        "prev_period": params["comparison_prev_period"],
    }

    tqdm.write(f"  → [{index}/{total}] {region_name}")
    region_rows: list[dict] = []
    all_ok = True

    for comp_key, comp_term_id in comp_map.items():
        tqdm.write(f"       {region_name}  [{comp_key}] fetching…")
        _set_active(region_name, comp_key)
        try:
            rows = fetch_tree(
                session, params,
                region_id=region_id,
                comparison_term_id=comp_term_id,
                parent_id="",
            )
            for r in rows:
                r["oblast"]          = region_name
                r["region"]          = region_name
                r["comparison_type"] = comp_key
            region_rows.extend(rows)
            tqdm.write(f"       {region_name}  [{comp_key}] done — {len(rows)} rows")
        except Exception as exc:
            tqdm.write(f"       {region_name}  [{comp_key}] ERROR: {exc}")
            all_ok = False

    _set_active(region_name, None)
    return region_name, region_rows, all_ok


def fetch_all(
    session: requests.Session,
    params: dict,
    checkpoint_path: Path | None = None,
    cookie: str = "",
    workers: int = FETCH_WORKERS,
) -> pd.DataFrame:
    """
    Fetch CPI for ALL oblasts × both comparison types using parallel workers.

    - *workers* regions are fetched concurrently (each gets its own HTTP session).
    - Progress is saved to *checkpoint_path* after every region completes so that
      a restart skips already-finished regions without re-fetching them.
    - Only regions where both comparison types succeeded are checkpointed.

    Returns DataFrame with columns:
        region, comparison_type, text, depth, parent_text, id, y{MMYYYY}…
    """
    if checkpoint_path is None:
        checkpoint_path = CHECKPOINT_CSV

    # ── Resume: load already-completed regions ──────────────────────────────
    checkpoint_df, done_regions = _load_checkpoint(checkpoint_path)

    pending = [r for r in CPI_REGIONS if r["name"] not in done_regions]
    total   = len(CPI_REGIONS)
    log.info("%d regions to fetch (%d already in checkpoint)",
             len(pending), len(done_regions))

    global _req_bar
    all_rows: list[dict] = []

    bar = tqdm(
        total=total,
        initial=len(done_regions),
        desc="Regions",
        unit="region",
        ncols=110,
        bar_format="{l_bar}{bar}| {n}/{total} [{elapsed}<{remaining}] {postfix}",
    )
    _req_bar = bar

    try:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = {
                pool.submit(_fetch_region, cookie, params, region, i + 1, total): region
                for i, region in enumerate(
                    sorted(pending, key=lambda r: CPI_REGIONS.index(r)), 0
                )
            }
            for future in as_completed(futures):
                if _stop_event.is_set():
                    break
                region_name, region_rows, region_ok = future.result()
                if region_ok and region_rows:
                    _append_checkpoint(checkpoint_path, region_rows)
                    bar.set_postfix_str(f"saved: {region_name}")
                elif not region_ok:
                    bar.set_postfix_str(f"ERRORS: {region_name}")
                    log.warning("  %s had errors — NOT checkpointed; will retry next run", region_name)
                bar.update(1)
                all_rows.extend(region_rows)
    except KeyboardInterrupt:
        _stop_event.set()
        tqdm.write("\n  [STOP] Ctrl+C received — waiting for in-flight requests to finish …")
    finally:
        _req_bar = None
        bar.close()
        if _stop_event.is_set():
            tqdm.write("  [STOP] Stopped. Progress saved to checkpoint. Run again to resume.")

    # ── Merge checkpoint rows + this-run rows ───────────────────────────────
    frames = []
    if checkpoint_df is not None and not checkpoint_df.empty:
        frames.append(checkpoint_df)
    if all_rows:
        frames.append(pd.DataFrame(all_rows))

    if not frames:
        log.error("No data was fetched!")
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)

    # Canonical column order
    key_cols = ["oblast", "region", "comparison_type", "depth", "parent_text", "text", "leaf", "id"]
    period_cols = sorted(
        [c for c in df.columns if re.match(r"^y\d{6}$", c)],
        key=lambda k: (int(k[3:]), int(k[1:3])),
    )
    ordered = [c for c in key_cols if c in df.columns] + period_cols
    df = df[[c for c in ordered if c in df.columns]]

    return df


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Fetch CPI from taldau.stat.gov.kz")
    parser.add_argument("--cookie",      default="",             help="Cookie string")
    parser.add_argument("--out",         default=str(DATA_CSV),  help="Output CSV path")
    parser.add_argument("--checkpoint",  default=str(CHECKPOINT_CSV), help="Checkpoint CSV path")
    parser.add_argument("--workers",     type=int, default=FETCH_WORKERS, help="Parallel region workers")
    parser.add_argument("--no-resume",   action="store_true",    help="Ignore checkpoint and fetch everything from scratch")
    args = parser.parse_args()

    cookie = args.cookie or _load_cookie()
    if not cookie:
        log.warning(
            "No cookie provided. Requests may fail with HTTP 302/401.\n"
            "Provide cookie via --cookie or place it in %s",
            COOKIE_FILE,
        )

    session = make_session(cookie)

    # Use hardcoded (pre-discovered) parameters
    params = dict(KNOWN_PARAMS)
    log.info("Using parameters: %s", params)
    log.info("Workers: %d", args.workers)

    # Verify cookie and connectivity before starting
    if not check_connection(session, params):
        tqdm.write("Aborting — fix your cookie and retry.")
        sys.exit(1)

    # Discover period names for nicer labels in output
    period_name_map, date_list, name_list = get_period_names(session, params)

    checkpoint_path = Path(args.checkpoint)
    if args.no_resume and checkpoint_path.exists():
        checkpoint_path.unlink()
        log.info("--no-resume: deleted existing checkpoint")

    # Fetch (parallel, resumes from checkpoint automatically unless --no-resume)
    df = fetch_all(session, params, checkpoint_path=checkpoint_path,
                   cookie=cookie, workers=args.workers)
    if df.empty:
        log.error("Nothing fetched — check cookie and network access.")
        sys.exit(1)

    out_path = Path(args.out)
    df.to_csv(out_path, index=False, encoding="utf-8-sig")
    log.info("Saved %d rows → %s", len(df), out_path)

    # Save params + period names for downstream scripts
    params_out = {
        **params,
        "date_list":       date_list,
        "period_name_list": name_list,
        "period_name_map": period_name_map,
        "target_periods":  TARGET_PERIODS,
    }
    params_path = out_path.parent / "cpi_params.json"
    params_path.write_text(
        json.dumps(params_out, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log.info("Saved params → %s", params_path)

    return df, params_out


if __name__ == "__main__":
    main()
