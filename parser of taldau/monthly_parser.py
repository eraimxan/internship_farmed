"""
Докачивает 11 недостающих районов и добавляет в MONTHLY_FINAL.csv

Запуск:
    python fetch_missing_districts.py

Результат: output/MONTHLY_FINAL.csv (обновлённый)
"""

import requests
import pandas as pd
import time
import os

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/123.0.0.0 Safari/537.36",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": "https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/701830",
    # ⚠️ ОБНОВИ КУКИ ПЕРЕД ЗАПУСКОМ!
    # F12 → Network → любой запрос к taldau → Headers → Cookie
    "Cookie": "_ym_uid=1758390801129425250; _ym_d=1758390801; ASP.NET_SessionId=su1ng2nr3ywtsmqt40n1kwf1"
})

BASE         = "https://taldau.stat.gov.kz/ru/NewIndex/GetIndexTreeData"
INDEX_ID     = "701830"
PERIOD_ID    = "8"           # Месяц с накоплением
DIC_IDS      = "68,61,3028"
TERMS_SUFFIX = "741926,741885"
TERM_ID      = "741885"
IDX          = "0"

TARGET_PERIODS = {
    "y012025": "Янв 2025",
    "y022025": "Янв-Фев 2025",
    "y032025": "Янв-Мар 2025",
    "y042025": "Янв-Апр 2025",
    "y052025": "Янв-Май 2025",
    "y062025": "Янв-Июн 2025",
    "y072025": "Янв-Июл 2025",
    "y082025": "Янв-Авг 2025",
    "y092025": "Янв-Сен 2025",
    "y102025": "Янв-Окт 2025",
    "y112025": "Янв-Ноя 2025",
    "y122025": "Янв-Дек 2025",
    "y012026": "Янв 2026",
    "y022026": "Янв-Фев 2026",
}

OUTPUT_DIR    = "output"
FINAL_CSV     = os.path.join(OUTPUT_DIR, "MONTHLY_FINAL.csv")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 11 недостающих районов с их ID
MISSING = [
    {"oblast": "АЛМАТИНСКАЯ ОБЛАСТЬ",          "name": "Г.А.АЛАТАУ",             "id": 77260375},
    {"oblast": "Г.АСТАНА",                      "name": "РАЙОН НҰРА",              "id": 77216130},
    {"oblast": "Г.АСТАНА",                      "name": "РАЙОН САРАЙШЫҚ",          "id": 77356522},
    {"oblast": "Г.ШЫМКЕНТ",                     "name": "РАЙОН ТУРАН",             "id": 77215757},
    {"oblast": "ТУРКЕСТАНСКАЯ ОБЛАСТЬ",         "name": "РАЙОН САУРАН",            "id": 77036054},
    {"oblast": "ВОСТОЧНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ","name": "РАЙОН САМАР",             "id": 77208143},
    {"oblast": "ВОСТОЧНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ","name": "РАЙОН МАРҚАКӨЛ",          "id": 77258303},
    {"oblast": "ВОСТОЧНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ","name": "РАЙОН ҮЛКЕН НАРЫН",       "id": 77258304},
    {"oblast": "ОБЛАСТЬ АБАЙ",                  "name": "РАЙОН АҚСУАТ",            "id": 77208142},
    {"oblast": "ОБЛАСТЬ АБАЙ",                  "name": "РАЙОН МАҚАНШЫ",           "id": 77258397},
    {"oblast": "ОБЛАСТЬ АБАЙ",                  "name": "РАЙОН ЖАҢАСЕМЕЙ",         "id": 77265110},
]


class CookieExpiredError(Exception):
    pass


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
        "idx":          IDX,
        "filter":       '[{"property":null,"value":null}]',
        "id":           parent_id if parent_id else TERM_ID,
    }
    for attempt in range(retries):
        try:
            r = session.post(BASE, data=data, timeout=90)
            time.sleep(delay)

            if r.status_code in (401, 403):
                raise CookieExpiredError(f"HTTP {r.status_code} — куки устарели")

            if r.status_code != 200:
                print(f"  ⚠️ HTTP {r.status_code}, попытка {attempt+1}/{retries}")
                time.sleep(delay * (attempt + 1))
                continue

            text = r.text.strip()
            if len(text) <= 2 or text in ("[]", "null", ""):
                if is_root:
                    raise CookieExpiredError("Корневой запрос вернул пустоту — куки устарели!")
                if attempt < retries - 1:
                    wait = delay * (attempt + 2)
                    print(f"  ⚠️ Пустой ответ (попытка {attempt+1}/{retries}), жду {wait:.1f}с...")
                    time.sleep(wait)
                    continue
                return []

            result = r.json()
            return result if isinstance(result, list) else []

        except CookieExpiredError:
            raise
        except requests.exceptions.ReadTimeout:
            wait = 10 * (attempt + 1)
            print(f"  ⚠️ Таймаут (попытка {attempt+1}/{retries}), жду {wait}с...")
            time.sleep(wait)
        except Exception as e:
            print(f"  ❌ Ошибка: {e}")
            time.sleep(delay * 2)
    return []


def fetch_recursive(p_terms, parent_id="", depth=0, parent_text="", max_depth=10):
    if depth > max_depth:
        return []
    is_root = (depth == 0 and parent_id == "")
    rows = fetch(p_terms, parent_id, is_root=is_root)
    all_rows = []
    for row in rows:
        enriched = dict(row)
        enriched["depth"] = depth
        enriched["parent_text"] = parent_text
        all_rows.append(enriched)
        if str(row.get("leaf", "true")).lower() == "false":
            children = fetch_recursive(
                p_terms,
                parent_id=str(row["id"]),
                depth=depth + 1,
                parent_text=row.get("text", ""),
                max_depth=max_depth
            )
            all_rows.extend(children)
    return all_rows


def check_session():
    """Тест на первом районе из списка"""
    item = MISSING[0]
    p_terms = f"{item['id']},{TERMS_SUFFIX}"
    print(f"Проверяем сессию ({item['name']}, id={item['id']})...")
    try:
        rows = fetch(p_terms, parent_id="", is_root=True)
    except CookieExpiredError:
        print("  ❌ Куки устарели!")
        return False
    if rows:
        print(f"  ✅ Сессия работает! Получено строк: {len(rows)}")
        # Показываем доступные колонки с данными
        sample = rows[0]
        y_cols = [k for k in sample.keys() if k.startswith('y')]
        print(f"  Колонки с данными: {y_cols[:10]}")
        return True
    print("  ❌ 0 строк — куки устарели!")
    return False


def main():
    print("=" * 60)
    print("Докачка 11 недостающих районов")
    print("=" * 60)

    # Проверяем сессию
    if not check_session():
        print("\nСТОП: обнови Cookie в скрипте и запусти снова")
        return

    # Загружаем текущий CSV
    if os.path.exists(FINAL_CSV):
        df_existing = pd.read_csv(FINAL_CSV, encoding="utf-8-sig", low_memory=False)
        print(f"\nЗагружен существующий CSV: {len(df_existing)} строк, {df_existing['region'].nunique()} районов")
    else:
        df_existing = pd.DataFrame()
        print("\nФайл MONTHLY_FINAL.csv не найден — создаём новый")

    new_rows = []

    try:
        for i, item in enumerate(MISSING, 1):
            print(f"\n[{i}/{len(MISSING)}] {item['oblast']} | {item['name']} (id={item['id']})...", end=" ", flush=True)

            p_terms = f"{item['id']},{TERMS_SUFFIX}"
            rows = fetch_recursive(p_terms, parent_id="")

            for row in rows:
                row["oblast"] = item["oblast"]
                row["region"] = item["name"]

            new_rows.extend(rows)
            print(f"✓ {len(rows)} строк")

    except CookieExpiredError as e:
        print(f"\n{'!'*60}")
        print(f"  🛑 СТОП: {e}")
        print(f"  Обнови Cookie и запусти снова.")
        print(f"{'!'*60}")
        if new_rows:
            print(f"  Успели собрать {len(new_rows)} строк — сохраняем частичный результат")
        else:
            return

    if not new_rows:
        print("\n⚠️ Новых данных не получено")
        return

    # Собираем колонки финального файла
    base_cols   = ["oblast", "region", "depth", "parent_text", "text", "leaf", "id"]
    period_cols = list(TARGET_PERIODS.keys())

    df_new = pd.DataFrame(new_rows)

    # Добавляем отсутствующие period колонки как None
    for col in period_cols:
        if col not in df_new.columns:
            df_new[col] = None

    # Пустая строка → None (нет данных на сайте)
    for col in period_cols:
        df_new[col] = df_new[col].replace("", None)
        df_new[col] = pd.to_numeric(df_new[col], errors='coerce')

    # Оставляем только нужные колонки
    keep_cols = [c for c in base_cols if c in df_new.columns] + \
                [c for c in period_cols if c in df_new.columns]
    df_new = df_new[keep_cols]

    print(f"\nНовых строк: {len(df_new)}")
    print(f"Новых районов: {df_new['region'].nunique()}")

    # Проверяем нет ли дублей (на случай повторного запуска)
    if not df_existing.empty:
        existing_regions = set(df_existing['region'].unique())
        new_regions = set(df_new['region'].unique())
        overlap = existing_regions & new_regions
        if overlap:
            print(f"  ⚠️ Удаляем дубли из существующего CSV: {overlap}")
            df_existing = df_existing[~df_existing['region'].isin(overlap)]

        # Выравниваем колонки
        for col in df_new.columns:
            if col not in df_existing.columns:
                df_existing[col] = None
        for col in df_existing.columns:
            if col not in df_new.columns:
                df_new[col] = None

        df_final = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_final = df_new

    # Сохраняем
    df_final.to_csv(FINAL_CSV, index=False, encoding="utf-8-sig")

    print(f"\n{'='*60}")
    print(f"✅ ГОТОВО!")
    print(f"   Итого строк:   {len(df_final)}")
    print(f"   Итого районов: {df_final['region'].nunique()}")
    print(f"   Файл:          {FINAL_CSV}")
    print(f"{'='*60}")

    # Показываем итоговое покрытие
    print("\nПокрытие по областям:")
    coverage = df_final.groupby('oblast')['region'].nunique().sort_index()
    for oblast, cnt in coverage.items():
        print(f"  {oblast}: {cnt} районов")


if __name__ == "__main__":
    main()