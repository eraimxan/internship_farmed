"""
core/parser.py — HTTP-парсинг данных с taldau.stat.gov.kz

Поддерживает:
  - Рекурсивный обход дерева регионов
  - Месячные и годовые данные (разные PERIOD_ID)
  - Retry-логику и проверку сессии
"""

import logging
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

from config.settings import (
    BASE_URL, INDEX_ID, TERM_ID,
    PERIOD_ID_MONTHLY, PERIOD_ID_YEARLY,
    MONTHLY_DIC_IDS, MONTHLY_TERMS_SUFFIX, MONTHLY_IDX,
    YEARLY_DIC_IDS, YEARLY_TERMS_SUFFIX, YEARLY_IDX,
    REQUEST_DELAY, RETRY_COUNT, REQUEST_TIMEOUT, MAX_DEPTH, MAX_WORKERS,
    PROGRESS_SAVE_INTERVAL,
)

logger = logging.getLogger(__name__)


def _get_params(period_id: str) -> tuple[str, str, str]:
    """Возвращает (dic_ids, terms_suffix, idx) для заданного period_id."""
    if period_id == PERIOD_ID_YEARLY:
        return YEARLY_DIC_IDS, YEARLY_TERMS_SUFFIX, YEARLY_IDX
    return MONTHLY_DIC_IDS, MONTHLY_TERMS_SUFFIX, MONTHLY_IDX


# ─────────────────────────────────────────────
# HTTP СЕССИЯ
# ─────────────────────────────────────────────

def make_session(cookie: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent":        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                             "AppleWebKit/537.36 Chrome/123.0.0.0 Safari/537.36",
        "Accept":            "application/json, text/javascript, */*; q=0.01",
        "Accept-Language":   "ru-RU,ru;q=0.9,en;q=0.8",
        "X-Requested-With":  "XMLHttpRequest",
        "Referer":           f"https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/{INDEX_ID}",
        "Cookie":            cookie,
    })
    return s


# ─────────────────────────────────────────────
# ПРОВЕРКА СЕССИИ
# ─────────────────────────────────────────────

def check_session_ok(cookie: str, period_id: str = PERIOD_ID_MONTHLY) -> tuple[bool, list[str]]:
    """
    Возвращает (True, список_доступных_периодов) если сессия активна,
    иначе (False, []).

    Args:
        cookie: строка Cookie
        period_id: "8" для месячных, "7" для годовых
    """
    dic_ids, terms_suffix, idx = _get_params(period_id)
    session = make_session(cookie)
    # Проверяем на первом районе — Кокшетау Г.А. (id=247784)
    data = {
        "p_parent_id":  "",
        "p_index_id":   INDEX_ID,
        "p_keyword":    "",
        "p_period_id":  period_id,
        "p_measure_id": "1",
        "p_term_id":    TERM_ID,
        "p_terms":      f"247784,{terms_suffix}",
        "p_dicIds":     dic_ids,
        "idx":          idx,
        "filter":       '[{"property":null,"value":null}]',
        "id":           TERM_ID,
    }
    try:
        r = session.post(BASE_URL, data=data, timeout=30)
        if r.status_code != 200:
            logger.warning(f"Session check: HTTP {r.status_code}")
            return False, []
        rows = r.json()
        if not rows:
            return False, []
        periods = sorted(
            [k for k in rows[0].keys() if re.match(r"^y\d{6}$", k)]
        )
        return True, periods
    except Exception as e:
        logger.error(f"Session check failed: {e}")
        return False, []


def check_session_both(cookie: str) -> tuple[bool, list[str], list[str]]:
    """
    Проверяет сессию и возвращает доступные периоды для обоих типов.

    Returns:
        (ok, monthly_periods, yearly_periods)
    """
    ok_m, monthly = check_session_ok(cookie, PERIOD_ID_MONTHLY)
    if not ok_m:
        return False, [], []
    ok_y, yearly = check_session_ok(cookie, PERIOD_ID_YEARLY)
    # Сессия валидна если хотя бы месячные работают
    return True, monthly, yearly if ok_y else []


# ─────────────────────────────────────────────
# ОЧИСТКА ЗНАЧЕНИЙ СТРОКИ
# ─────────────────────────────────────────────

def _clean_row(row: dict) -> dict:
    cleaned: dict = {}
    for k, v in row.items():
        if re.match(r"^y\d{6}$", k) or re.match(r"^\d+$", k):
            if v == "" or v is None:
                cleaned[k] = None
            else:
                try:
                    cleaned[k] = int(v)
                except (ValueError, TypeError):
                    try:
                        cleaned[k] = float(v)
                    except Exception:
                        cleaned[k] = v
        else:
            cleaned[k] = v
    return cleaned


# ─────────────────────────────────────────────
# РЕКУРСИВНЫЙ ПАРСИНГ
# ─────────────────────────────────────────────

def fetch_recursive(
    session: requests.Session,
    p_terms: str,
    period_id: str = PERIOD_ID_MONTHLY,
    parent_id: str = "",
    depth: int = 0,
    parent_text: str = "",
) -> list[dict]:
    """
    Рекурсивно обходит дерево данных для заданного p_terms.

    Args:
        session: HTTP-сессия с Cookie
        p_terms: строка p_terms для запроса
        period_id: "8" для месячных, "7" для годовых
        parent_id: ID родительского узла
        depth: текущая глубина рекурсии
        parent_text: текст родительского узла

    Returns:
        Список строк (dict) со всеми периодами.
    """
    if depth > MAX_DEPTH:
        return []

    dic_ids, _, idx = _get_params(period_id)

    data = {
        "p_parent_id":  parent_id,
        "p_index_id":   INDEX_ID,
        "p_keyword":    "",
        "p_period_id":  period_id,
        "p_measure_id": "1",
        "p_term_id":    TERM_ID,
        "p_terms":      p_terms,
        "p_dicIds":     dic_ids,
        "idx":          idx,
        "filter":       '[{"property":null,"value":null}]',
        "id":           parent_id if parent_id else TERM_ID,
    }

    for attempt in range(RETRY_COUNT):
        try:
            r = session.post(BASE_URL, data=data, timeout=REQUEST_TIMEOUT)
            time.sleep(REQUEST_DELAY)

            if r.status_code != 200:
                logger.warning(f"HTTP {r.status_code} depth={depth} attempt={attempt}")
                time.sleep(3 * (attempt + 1))
                continue

            text = r.text.strip()
            if len(text) <= 2 or text in ("[]", "null", ""):
                if attempt < RETRY_COUNT - 1:
                    time.sleep(4 * (attempt + 1))
                    continue
                return []

            rows = r.json()
            if not isinstance(rows, list):
                return []

            all_rows: list[dict] = []
            for row in rows:
                enriched = _clean_row(dict(row))
                enriched["depth"] = depth
                enriched["parent_text"] = parent_text
                all_rows.append(enriched)

                if str(row.get("leaf", "true")).lower() == "false":
                    all_rows.extend(
                        fetch_recursive(
                            session,
                            p_terms,
                            period_id=period_id,
                            parent_id=str(row["id"]),
                            depth=depth + 1,
                            parent_text=row.get("text", ""),
                        )
                    )

            return all_rows

        except Exception as e:
            logger.warning(f"Fetch depth={depth} attempt={attempt}: {e}")
            time.sleep(5)

    return []


# ─────────────────────────────────────────────
# ПАРСИНГ ВСЕХ РЕГИОНОВ
# ─────────────────────────────────────────────

def parse_all_regions(
    cookie: str,
    new_cols: list[str],
    all_regions: dict,
    period_id: str = PERIOD_ID_MONTHLY,
    progress_path=None,
) -> tuple[list[dict], list[str]]:
    """
    Парсит все регионы параллельно (MAX_WORKERS воркеров) и возвращает:
      - all_rows: список строк данных
      - errors: список ошибок по регионам
    """
    import pandas as pd
    from pathlib import Path

    all_rows: list[dict] = []
    errors: list[str] = []
    csv_lock = threading.Lock()

    # ── Resume ────────────────────────────────────────────────────
    done_regions: set[tuple[str, str]] = set()

    if progress_path and Path(progress_path).exists():
        try:
            prev = pd.read_csv(progress_path, encoding="utf-8-sig", low_memory=False)
            if not prev.empty and "oblast" in prev.columns and "region" in prev.columns:
                all_rows = prev.to_dict("records")
                done_regions = set(
                    zip(prev["oblast"].astype(str), prev["region"].astype(str))
                )
                logger.info(
                    f"Resume: загружено {len(all_rows):,} строк, "
                    f"{len(done_regions)} районов пройдено"
                )
        except Exception as e:
            logger.warning(f"Resume: не удалось загрузить прогресс: {e}")
            all_rows = []
            done_regions = set()

    # ── Формируем очередь задач ───────────────────────────────────
    tasks = [
        (oblast_name, item)
        for oblast_name, districts in all_regions.items()
        for item in districts
        if (oblast_name, item["name"]) not in done_regions
    ]
    total = sum(len(v) for v in all_regions.values())
    done_count = len(done_regions)

    _, terms_suffix, _ = _get_params(period_id)

    # One persistent session per worker thread (avoids TCP handshake per district)
    thread_local = threading.local()

    def get_session() -> requests.Session:
        if not hasattr(thread_local, "session"):
            thread_local.session = make_session(cookie)
        return thread_local.session

    def parse_district(oblast_name: str, item: dict) -> tuple[str, str, list[dict]]:
        p_terms = f"{item['id']},{terms_suffix}"
        rows = fetch_recursive(get_session(), p_terms, period_id=period_id, parent_id="")
        return oblast_name, item["name"], rows

    # ── Параллельный парсинг ──────────────────────────────────────
    completed_count = 0

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(parse_district, oblast_name, item): (oblast_name, item)
            for oblast_name, item in tasks
        }

        for future in as_completed(futures):
            oblast_name, item = futures[future]
            try:
                oblast_name, region_name, rows = future.result()

                if not rows:
                    logger.warning(f"⚠ {oblast_name} / {region_name} — 0 строк (пропуск)")
                    continue

                new_rows = []
                for r in rows:
                    r["oblast"] = oblast_name
                    r["region"] = region_name
                    keep = {
                        k: v for k, v in r.items()
                        if k in ("oblast", "region", "depth", "parent_text", "text", "leaf", "id")
                        or k in new_cols
                    }
                    new_rows.append(keep)

                with csv_lock:
                    all_rows.extend(new_rows)
                    completed_count += 1
                    logger.info(f"[{done_count + completed_count}/{total}] ✓ {oblast_name} / {region_name} — {len(rows)} строк")

                    if progress_path and all_rows and completed_count % PROGRESS_SAVE_INTERVAL == 0:
                        try:
                            pd.DataFrame(all_rows).to_csv(
                                progress_path, index=False, encoding="utf-8-sig"
                            )
                        except Exception:
                            pass

            except Exception as e:
                msg = f"{oblast_name}/{item['name']}: {e}"
                with csv_lock:
                    errors.append(msg)
                logger.error(msg)

    # Final save to capture any districts not on a save boundary
    if progress_path and all_rows:
        try:
            pd.DataFrame(all_rows).to_csv(progress_path, index=False, encoding="utf-8-sig")
        except Exception:
            pass

    return all_rows, errors
