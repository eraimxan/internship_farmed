"""
patch_missing.py
================
Дозагружает 4 недостающих района с taldau.stat.gov.kz и добавляет их
в ALL_KAZAKHSTAN_PATCHED.csv, сохраняя итог в ALL_KAZAKHSTAN_FULL.csv.

Недостающие районы (выявлены автоматически):
  - АЛМАТИНСКАЯ ОБЛАСТЬ  ->  Г.А.АЛАТАУ        (id=77260375)
  - Г.АСТАНА             ->  РАЙОН НУРА         (id=77216130)
  - Г.АСТАНА             ->  РАЙОН САРАЙШЫК     (id=77356522)
  - Г.ШЫМКЕНТ            ->  РАЙОН ТУРАН        (id=77215757)

КАК ЗАПУСТИТЬ:
  1. Открой https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/701830
  2. F12 -> Network -> любой XHR -> Request Headers -> скопируй Cookie
  3. Вставь в переменную COOKIE ниже (заменив старое значение)
  4. Положи ALL_KAZAKHSTAN_PATCHED.csv рядом со скриптом
  5. python patch_missing.py
"""

import requests
import pandas as pd
import time
from datetime import datetime
import os
import sys

# -------------------------------------------------------
# ОБНОВИ ПЕРЕД ЗАПУСКОМ!
# -------------------------------------------------------
COOKIE = "_ym_uid=1758390801129425250; _ym_d=1758390801; ASP.NET_SessionId=su1ng2nr3ywtsmqt40n1kwf1"

INPUT_CSV  = "output/ALL_KAZAKHSTAN_PATCHED.csv"
OUTPUT_CSV = "ALL_KAZAKHSTAN_FULL.csv"

# -------------------------------------------------------
# ПАРАМЕТРЫ API (не меняй)
# -------------------------------------------------------
BASE         = "https://taldau.stat.gov.kz/ru/NewIndex/GetIndexTreeData"
INDEX_ID     = "701830"
PERIOD_ID    = "7"
DIC_IDS      = "68,776,59,3028,4043"
TERM_ID      = "741885"
TERMS_SUFFIX = "741917,741907,741885,19202525"

# -------------------------------------------------------
# НЕДОСТАЮЩИЕ РАЙОНЫ
# -------------------------------------------------------
MISSING_REGIONS = [
    {"oblast": "АЛМАТИНСКАЯ ОБЛАСТЬ", "id": 77260375, "name": "Г.А.АЛАТАУ"},
    {"oblast": "Г.АСТАНА",            "id": 77216130, "name": "РАЙОН НУРА"},
    {"oblast": "Г.АСТАНА",            "id": 77356522, "name": "РАЙОН САРАЙШЫК"},
    {"oblast": "Г.ШЫМКЕНТ",           "id": 77215757, "name": "РАЙОН ТУРАН"},
]

# -------------------------------------------------------
session = requests.Session()
session.headers.update({
    "User-Agent":       "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/123.0.0.0 Safari/537.36",
    "Accept":           "application/json, text/javascript, */*; q=0.01",
    "Accept-Language":  "ru-RU,ru;q=0.9,en;q=0.8",
    "X-Requested-With": "XMLHttpRequest",
    "Referer":          "https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/701830",
    "Cookie":           COOKIE,
})

stats = {"requests": 0, "empty": 0, "errors": 0}


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


class CookieExpiredError(Exception):
    pass


def clean_row(row):
    cleaned = {}
    for k, v in row.items():
        is_data = (
            (k.startswith("y") and k[1:].isdigit()) or
            k.isdigit() or
            (len(k) >= 2 and k[:2].isdigit())
        )
        if is_data:
            if v == "" or v is None:
                cleaned[k] = None
            else:
                try:
                    cleaned[k] = int(v)
                except (ValueError, TypeError):
                    try:
                        cleaned[k] = float(v)
                    except (ValueError, TypeError):
                        cleaned[k] = v
        else:
            cleaned[k] = v
    return cleaned


def fetch(p_terms, parent_id="", retries=5, delay=1.5, is_root=False):
    data = {
        "p_parent_id":  parent_id,
        "p_index_id":   INDEX_ID,
        "p_keyword":    "",
        "p_period_id":  PERIOD_ID,
        "p_measure_id": "1",
        "p_term_id":    TERM_ID,
        "p_terms":      p_terms,
        "p_dicIds":     DIC_IDS,
        "idx":          "3",
        "filter":       '[{"property":null,"value":null}]',
        "id":           parent_id if parent_id else TERM_ID,
    }
    stats["requests"] += 1

    for attempt in range(retries):
        try:
            r = session.post(BASE, data=data, timeout=90)
            time.sleep(delay)

            if r.status_code in (401, 403):
                raise CookieExpiredError(f"HTTP {r.status_code}")

            if r.status_code != 200:
                log(f"  WARNING HTTP {r.status_code}, попытка {attempt+1}/{retries}")
                time.sleep(delay * (attempt + 1))
                continue

            text = r.text.strip()
            if len(text) <= 2 or text in ("[]", "null", ""):
                stats["empty"] += 1
                if is_root:
                    raise CookieExpiredError("Корневой узел пуст — куки устарели!")
                if attempt < retries - 1:
                    wait = delay * (attempt + 2)
                    log(f"  WARNING Пустой ответ (попытка {attempt+1}/{retries}), жду {wait:.0f}с...")
                    time.sleep(wait)
                    continue
                return []

            result = r.json()
            return result if isinstance(result, list) else []

        except CookieExpiredError:
            raise
        except requests.exceptions.ReadTimeout:
            wait = 10 * (attempt + 1)
            log(f"  WARNING Таймаут (попытка {attempt+1}/{retries}), жду {wait}с...")
            time.sleep(wait)
        except Exception as e:
            stats["errors"] += 1
            log(f"  ERROR {e}")
            time.sleep(delay * 2)

    return []


def fetch_recursive(p_terms, parent_id="", depth=0, parent_text="", max_depth=10):
    if depth > max_depth:
        return []

    is_root = (depth == 0 and parent_id == "")
    rows = fetch(p_terms, parent_id, is_root=is_root)
    all_rows = []

    for row in rows:
        enriched = clean_row(dict(row))
        enriched["depth"] = depth
        enriched["parent_text"] = parent_text
        all_rows.append(enriched)

        if str(row.get("leaf", "true")).lower() == "false":
            children = fetch_recursive(
                p_terms,
                parent_id=str(row["id"]),
                depth=depth + 1,
                parent_text=row.get("text", ""),
                max_depth=max_depth,
            )
            all_rows.extend(children)

    return all_rows


def check_session():
    log("Проверяем сессию (тест: РАЙОН ТУРАН id=77215757)...")
    try:
        rows = fetch(f"77215757,{TERMS_SUFFIX}", parent_id="", is_root=True)
    except CookieExpiredError:
        log("  FAIL Куки устарели!")
        return False
    if rows:
        log(f"  OK Сессия жива! Строк в тесте: {len(rows)}")
        return True
    log("  FAIL 0 строк — куки устарели!")
    return False


def main():
    # 1. Проверка куки
    if not check_session():
        print("\n" + "="*60)
        print("  СТОП: куки не работают.")
        print("  1. Зайди на https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/701830")
        print("  2. F12 -> Network -> любой XHR -> Request Headers -> скопируй Cookie")
        print("  3. Вставь в переменную COOKIE в начале этого файла")
        print("="*60)
        sys.exit(1)

    # 2. Загрузка исходного CSV
    if not os.path.exists(INPUT_CSV):
        print(f"FAIL Файл не найден: {INPUT_CSV}")
        print(f"   Положи его рядом со скриптом и запусти снова.")
        sys.exit(1)

    log(f"Загружаем {INPUT_CSV}...")
    df_existing = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)
    log(f"  Строк: {len(df_existing)}, районов: {df_existing['region'].nunique()}, "
        f"областей: {df_existing['oblast'].nunique()}")

    # 3. Дозагрузка недостающих районов
    new_rows_all = []

    for item in MISSING_REGIONS:
        oblast = item["oblast"]
        region_id = item["id"]
        name = item["name"]
        log(f"\n{'--'*25}")
        log(f"Фетчим: {oblast}  ->  {name}  (id={region_id})")

        p_terms = f"{region_id},{TERMS_SUFFIX}"
        try:
            rows = fetch_recursive(p_terms, parent_id="")
        except CookieExpiredError:
            log("  STOP Куки протухли в процессе! Сохраняем что успели...")
            break

        if not rows:
            log(f"  WARNING ПУСТО — район {name} не вернул данных. Пропускаем.")
            continue

        for row in rows:
            row["oblast"] = oblast
            row["region"] = name

        new_rows_all.extend(rows)
        log(f"  OK {len(rows)} строк для {name}")

    # 4. Объединение и сохранение
    if not new_rows_all:
        log("\nFAIL Новых данных нет. Ничего не сохраняем.")
        sys.exit(1)

    df_new = pd.DataFrame(new_rows_all)
    log(f"\nНовых строк собрано: {len(df_new)}")

    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    log(f"\n{'='*60}")
    log(f"ГОТОВО!")
    log(f"   Строк в исходнике:  {len(df_existing)}")
    log(f"   Новых строк:        {len(df_new)}")
    log(f"   Итого строк:        {len(df_combined)}")
    log(f"   Районов итого:      {df_combined['region'].nunique()}")
    log(f"   Областей итого:     {df_combined['oblast'].nunique()}")
    log(f"   Запросов к API:     {stats['requests']}")
    log(f"   Сохранено в:        {OUTPUT_CSV}")
    log(f"{'='*60}")

    # 5. Итоговая сводка по районам
    missing_names = {x["name"] for x in MISSING_REGIONS}
    print("\n=== ИТОГОВЫЕ РАЙОНЫ ПО ОБЛАСТЯМ ===")
    for obl in sorted(df_combined["oblast"].unique()):
        regions = sorted(df_combined[df_combined["oblast"] == obl]["region"].unique())
        print(f"\n  {obl}  ({len(regions)} районов):")
        for r in regions:
            marker = " <- NEW" if r in missing_names else ""
            print(f"    - {r}{marker}")


if __name__ == "__main__":
    main()