"""
core/storage.py — Управление мастер-файлом, состоянием и cookies.

Принцип: при обновлении новые колонки ДОПИСЫВАЮТСЯ.
Пустые ячейки в существующих строках заполняются новыми данными.
Непустые ячейки НЕ перезаписываются (данные не теряются).
"""

import json
import logging
import re
from datetime import datetime

import pandas as pd

from config.settings import (
    MASTER_CSV, STATE_FILE, COOKIE_FILE, KEY_COLS, BASE_DIR,
)
from core.periods import period_sort_key, is_yearly_key, is_monthly_key, is_period_key, fmt_period

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# STATE
# ─────────────────────────────────────────────

def load_state() -> dict:
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {
        "loaded_periods": [],
        "loaded_yearly": [],
        "available_periods": [],
        "available_yearly": [],
        "last_check": None,
        "last_update": None,
    }


def save_state(state: dict) -> None:
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────
# COOKIE
# ─────────────────────────────────────────────

def get_cookie() -> str:
    return COOKIE_FILE.read_text(encoding="utf-8").strip() if COOKIE_FILE.exists() else ""


def save_cookie(cookie: str) -> None:
    COOKIE_FILE.write_text(cookie.strip(), encoding="utf-8")


# ─────────────────────────────────────────────
# МАСТЕР-ФАЙЛ
# ─────────────────────────────────────────────

def load_master() -> pd.DataFrame:
    if MASTER_CSV.exists():
        df = pd.read_csv(MASTER_CSV, encoding="utf-8-sig", low_memory=False)
        logger.info(f"Мастер загружен: {len(df):,} строк, {len(df.columns)} колонок")
        return df
    logger.info("Мастер-файл не найден — возвращаю пустой DataFrame")
    return pd.DataFrame()


def save_master(df: pd.DataFrame) -> None:
    df.to_csv(MASTER_CSV, index=False, encoding="utf-8-sig")
    logger.info(f"Мастер сохранён: {len(df):,} строк, {len(df.columns)} колонок")


# ─────────────────────────────────────────────
# СЛИЯНИЕ: ДОПИСЫВАЕМ НОВЫЕ КОЛОНКИ, НЕ ТРОГАЕМ СТАРЫЕ
# ─────────────────────────────────────────────

def merge_new_data(
    master: pd.DataFrame,
    new_df: pd.DataFrame,
    new_period_cols: list[str],
) -> pd.DataFrame:
    """
    Добавляет только новые колонки периодов в мастер-файл.

    Алгоритм:
      - Строки, которые уже есть в мастере → дописываем только ПУСТЫЕ ячейки
      - Строки, которых нет в мастере → добавляем целиком
      - Непустые существующие данные НЕ изменяются
    """
    if master.empty:
        logger.info("Мастер пустой — инициализируем с новыми данными")
        return new_df.copy()

    master = master.copy()

    # Добавляем новые колонки в мастер, если ещё не существуют
    for col in new_period_cols:
        if col not in master.columns:
            master[col] = None

    # Гарантируем наличие ключевых полей
    for col in KEY_COLS:
        if col not in master.columns:
            master[col] = None
        if col not in new_df.columns:
            new_df = new_df.copy()
            new_df[col] = None

    master["_key"] = master[KEY_COLS].astype(str).agg("|".join, axis=1)
    new_df = new_df.copy()
    new_df["_key"] = new_df[KEY_COLS].astype(str).agg("|".join, axis=1)

    key_to_idx = {k: i for i, k in enumerate(master["_key"])}

    updated, added, skipped = 0, 0, 0
    new_rows_to_add = []

    for _, new_row in new_df.iterrows():
        key = new_row["_key"]
        if key in key_to_idx:
            idx = key_to_idx[key]
            for col in new_period_cols:
                if col in new_row.index and pd.notna(new_row[col]):
                    # Записываем ТОЛЬКО если текущая ячейка пустая
                    if pd.isna(master.at[idx, col]) or master.at[idx, col] is None:
                        master.at[idx, col] = new_row[col]
                        updated += 1
                    else:
                        skipped += 1
        else:
            new_rows_to_add.append(new_row)
            added += 1

    if new_rows_to_add:
        new_block = pd.DataFrame(new_rows_to_add)
        for col in master.columns:
            if col not in new_block.columns:
                new_block[col] = None
        master = pd.concat([master, new_block[master.columns]], ignore_index=True)

    master.drop(columns=["_key"], inplace=True)
    logger.info(f"Merge: обновлено {updated} ячеек, добавлено {added} строк, пропущено {skipped} непустых")
    return master


# ─────────────────────────────────────────────
# ИМПОРТ СУЩЕСТВУЮЩИХ CSV (одноразовая операция)
# ─────────────────────────────────────────────

def import_existing_csvs(full_path: str, monthly_path: str) -> pd.DataFrame:
    """
    Импортирует ALL_KAZAKHSTAN_FULL.csv + MONTHLY_FINAL.csv в мастер-файл.
    Вызывается один раз при первоначальной инициализации.
    """
    logger.info("Импорт существующих CSV...")

    df_full = pd.read_csv(full_path, encoding="utf-8-sig", low_memory=False)
    # Автоматически определяем годовые колонки (y12XXXX)
    year_cols = [c for c in df_full.columns if re.match(r"^y12\d{4}$", c)]
    base = ["id", "text", "leaf", "depth", "parent_text", "oblast", "region"]
    df_full = df_full[[c for c in base if c in df_full.columns] + year_cols]

    # Дедупликация по ключевым полям
    dups = df_full.duplicated(subset=KEY_COLS, keep=False)
    if dups.sum() > 0:
        seen: dict = {}
        for _, row in df_full[dups].iterrows():
            k = tuple(row[c] for c in KEY_COLS)
            if k not in seen:
                seen[k] = row.copy()
            else:
                for col in year_cols:
                    if pd.isna(seen[k][col]) and pd.notna(row[col]):
                        seen[k][col] = row[col]
                if pd.notna(row["id"]) and pd.notna(seen[k]["id"]):
                    seen[k]["id"] = min(seen[k]["id"], row["id"])
        df_full = pd.concat(
            [df_full[~dups], pd.DataFrame(list(seen.values()))],
            ignore_index=True
        )

    df_monthly = pd.read_csv(monthly_path, encoding="utf-8-sig", low_memory=False)
    monthly_cols = [c for c in df_monthly.columns if is_period_key(c) and not is_yearly_key(c)]
    df_monthly = df_monthly[[c for c in base if c in df_monthly.columns] + monthly_cols]

    df = pd.merge(df_full, df_monthly, on=KEY_COLS, how="outer", suffixes=("_f", "_m"))
    for field in ["id", "parent_text", "leaf"]:
        if f"{field}_f" in df.columns:
            df[field] = df[f"{field}_f"].combine_first(df.get(f"{field}_m", pd.Series()))
            df.drop(columns=[f"{field}_f", f"{field}_m"], errors="ignore", inplace=True)

    df["depth"] = df["depth"].fillna(0).astype(int)
    df["id"] = pd.to_numeric(df["id"], errors="coerce").astype("Int64")
    logger.info(f"Импорт завершён: {len(df):,} строк")
    return df


# ─────────────────────────────────────────────
# ВСПОМОГАТЕЛЬНЫЕ
# ─────────────────────────────────────────────

def get_full_period_map(df: pd.DataFrame) -> dict[str, str]:
    """
    Возвращает все периоды из мастера в хронологическом порядке.
    Автоматически определяет как годовые, так и месячные колонки.
    """
    from core.periods import period_label, parse_period_key, yearly_label

    period_map: dict[str, str] = {}

    # Собираем ВСЕ колонки периодов из DataFrame
    all_period_cols = sorted(
        [c for c in df.columns if is_period_key(c)],
        key=period_sort_key,
    )

    for c in all_period_cols:
        parsed = parse_period_key(c)
        if parsed:
            month, year = parsed
            # period_label уже корректно обрабатывает month==12 → "YYYY (год)"
            period_map[c] = period_label(month, year)

    return period_map


def mark_periods_loaded(new_cols: list[str], data_type: str = "monthly") -> None:
    """
    Добавляет новые периоды в state.json.

    Args:
        new_cols: список новых ключей периодов
        data_type: "monthly" или "yearly"
    """
    state = load_state()

    if data_type == "yearly":
        loaded = set(state.get("loaded_yearly", []))
        for p in new_cols:
            loaded.add(p)
        state["loaded_yearly"] = sorted(loaded, key=period_sort_key)
    else:
        loaded = set(state.get("loaded_periods", []))
        for p in new_cols:
            loaded.add(p)
        state["loaded_periods"] = sorted(loaded, key=period_sort_key)

    state["last_update"] = datetime.now().isoformat()
    save_state(state)


def get_all_loaded_periods() -> set[str]:
    """Возвращает все загруженные периоды (и месячные, и годовые)."""
    state = load_state()
    loaded = set(state.get("loaded_periods", []))
    loaded.update(state.get("loaded_yearly", []))
    return loaded


def get_master_period_cols() -> set[str]:
    """
    Returns period column keys that actually exist in master.csv.
    This is the authoritative source — state.json may drift out of sync.
    """
    if not MASTER_CSV.exists():
        return set()
    try:
        cols = pd.read_csv(MASTER_CSV, nrows=0, encoding="utf-8-sig").columns.tolist()
        return {c for c in cols if is_period_key(c)}
    except Exception:
        return set()
