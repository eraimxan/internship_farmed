"""
parse_monthly.py
Парсит помесячные данные (нарастающим итогом) по ВСЕМ регионам и районам Казахстана.
Логика полностью аналогична parse_yearly.py, но с параметрами для месячных данных.

Запуск:
    python parse_monthly.py

Результат:
    output_monthly/ALL_KAZAKHSTAN_MONTHLY.csv  — финальный файл
    output_monthly/<oblast>.csv                — по каждой области
    output_monthly/monthly_progress.csv        — чекпоинт
    output_monthly/done.txt                    — пройденные районы
    output_monthly/scrape_log.txt              — лог
"""

import requests
import pandas as pd
import time
import os
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────
# КОНФИГУРАЦИЯ
# ─────────────────────────────────────────────────────────────────────

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
PERIOD_ID    = "8"           # Месяц с накоплением (в отличие от "7" у годового)
DIC_IDS      = "68,61,3028"
TERMS_SUFFIX = "741926,741885"
TERM_ID      = "741885"
IDX          = "0"

# Целевые периоды — используются для очистки колонок в итоговом CSV
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

OUTPUT_DIR    = "output_monthly"
LOG_FILE      = os.path.join(OUTPUT_DIR, "scrape_log.txt")
PROGRESS_FILE = os.path.join(OUTPUT_DIR, "monthly_progress.csv")
DONE_FILE     = os.path.join(OUTPUT_DIR, "done.txt")
FINAL_CSV     = os.path.join(OUTPUT_DIR, "ALL_KAZAKHSTAN_MONTHLY.csv")

os.makedirs(OUTPUT_DIR, exist_ok=True)

stats = {"total_requests": 0, "empty_responses": 0, "errors": 0, "timeout": 0}

# ─────────────────────────────────────────────────────────────────────
# ЛОГИРОВАНИЕ
# ─────────────────────────────────────────────────────────────────────

def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

# ─────────────────────────────────────────────────────────────────────
# ОЧИСТКА СТРОК
# Сайт возвращает колонки вида: "y012025" (нарастающий итог за период)
# Пустая строка "" = данных нет; число = данные есть
# ─────────────────────────────────────────────────────────────────────

def clean_row(row):
    cleaned = {}
    for k, v in row.items():
        is_data_col = (
            (k.startswith("y") and k[1:].isdigit()) or
            k.isdigit() or
            (len(k) >= 2 and k[:2].isdigit())
        )
        if is_data_col:
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

# ─────────────────────────────────────────────────────────────────────
# ЗАПРОС
# ─────────────────────────────────────────────────────────────────────

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
    stats["total_requests"] += 1

    for attempt in range(retries):
        try:
            r = session.post(BASE, data=data, timeout=90)
            time.sleep(delay)

            if r.status_code in (401, 403):
                raise CookieExpiredError(f"HTTP {r.status_code} — куки устарели")

            if r.status_code != 200:
                log(f"  ⚠️ HTTP {r.status_code}, попытка {attempt+1}/{retries}")
                time.sleep(delay * (attempt + 1))
                continue

            text = r.text.strip()

            if len(text) <= 2 or text in ("[]", "null", ""):
                stats["empty_responses"] += 1
                if is_root:
                    raise CookieExpiredError("Корневой запрос вернул пустоту — куки устарели!")
                if attempt < retries - 1:
                    wait = delay * (attempt + 2)
                    log(f"  ⚠️ Пустой ответ (попытка {attempt+1}/{retries}), жду {wait:.1f}с...")
                    time.sleep(wait)
                    continue
                return []

            result = r.json()
            return result if isinstance(result, list) else []

        except CookieExpiredError:
            raise
        except requests.exceptions.ReadTimeout:
            stats["timeout"] += 1
            wait = 10 * (attempt + 1)
            log(f"  ⚠️ Таймаут (попытка {attempt+1}/{retries}), жду {wait}с...")
            time.sleep(wait)
        except Exception as e:
            stats["errors"] += 1
            log(f"  ❌ Ошибка: {e}")
            time.sleep(delay * 2)
    return []

# ─────────────────────────────────────────────────────────────────────
# РЕКУРСИВНЫЙ ОБХОД
# ─────────────────────────────────────────────────────────────────────

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
                max_depth=max_depth
            )
            all_rows.extend(children)

    return all_rows

# ─────────────────────────────────────────────────────────────────────
# ПРОВЕРКА СЕССИИ (первый район из АКМОЛИНСКОЙ ОБЛАСТИ)
# ─────────────────────────────────────────────────────────────────────

def check_session():
    log("Проверяем сессию (КОКШЕТАУ Г.А., id=247784)...")
    try:
        rows = fetch(f"247784,{TERMS_SUFFIX}", parent_id="", is_root=True)
    except CookieExpiredError:
        log("  ❌ Куки устарели!")
        return False
    if rows:
        log(f"  ✅ Сессия работает! Строк: {len(rows)}")
        y_cols = [k for k in rows[0].keys() if k.startswith("y")]
        log(f"  Доступные периоды: {y_cols[:14]}")
        return True
    log("  ❌ 0 строк — куки устарели!")
    return False

# ─────────────────────────────────────────────────────────────────────
# ВСЕ РЕГИОНЫ КАЗАХСТАНА (полный список из parse_yearly.py)
# ─────────────────────────────────────────────────────────────────────

ALL_REGIONS = {
    "АЛМАТИНСКАЯ ОБЛАСТЬ": [
        {"id": 250517,   "name": "ҚОНАЕВ Г.А."},
        {"id": 250736,   "name": "БАЛХАШСКИЙ РАЙОН"},
        {"id": 251043,   "name": "ЕНБЕКШИКАЗАХСКИЙ РАЙОН"},
        {"id": 251178,   "name": "ЖАМБЫЛСКИЙ РАЙОН"},
        {"id": 251556,   "name": "КАРАСАЙСКИЙ РАЙОН"},
        {"id": 251751,   "name": "РАЙЫМБЕКСКИЙ РАЙОН"},
        {"id": 251976,   "name": "ТАЛГАРСКИЙ РАЙОН"},
        {"id": 252156,   "name": "УЙГУРСКИЙ РАЙОН"},
        {"id": 252213,   "name": "ИЛИЙСКИЙ РАЙОН"},
        {"id": 20169323, "name": "КЕГЕНСКИЙ РАЙОН"},
        {"id": 77260375, "name": "Г.А.АЛАТАУ"},
    ],
    "АКМОЛИНСКАЯ ОБЛАСТЬ": [
        {"id": 247784, "name": "КОКШЕТАУ Г.А."},
        {"id": 247791, "name": "СТЕПНОГОРСК Г.А."},
        {"id": 247807, "name": "АККОЛЬСКИЙ РАЙОН"},
        {"id": 247855, "name": "АРШАЛЫНСКИЙ РАЙОН"},
        {"id": 247904, "name": "АСТРАХАНСКИЙ РАЙОН"},
        {"id": 247961, "name": "АТБАСАРСКИЙ РАЙОН"},
        {"id": 248026, "name": "БУЛАНДЫНСКИЙ РАЙОН"},
        {"id": 248076, "name": "ЕГИНДЫКОЛЬСКИЙ РАЙОН"},
        {"id": 248100, "name": "РАЙОН БИРЖАН САЛ"},
        {"id": 248165, "name": "ЕРЕЙМЕНТАУСКИЙ РАЙОН"},
        {"id": 248261, "name": "ЕСИЛЬСКИЙ РАЙОН"},
        {"id": 248322, "name": "ЖАКСЫНСКИЙ РАЙОН"},
        {"id": 248378, "name": "ЖАРКАИНСКИЙ РАЙОН"},
        {"id": 248433, "name": "ЗЕРЕНДИНСКИЙ РАЙОН"},
        {"id": 248558, "name": "КОРГАЛЖЫНСКИЙ РАЙОН"},
        {"id": 248592, "name": "САНДЫКТАУСКИЙ РАЙОН"},
        {"id": 248655, "name": "ЦЕЛИНОГРАДСКИЙ РАЙОН"},
        {"id": 248729, "name": "ШОРТАНДИНСКИЙ РАЙОН"},
        {"id": 248775, "name": "БУРАБАЙСКИЙ РАЙОН"},
        {"id": 77123369, "name": "КОСШЫ Г.А."},
    ],
    "АКТЮБИНСКАЯ ОБЛАСТЬ": [
        {"id": 248876, "name": "АКТОБЕ Г.А."},
        {"id": 248919, "name": "АЛГИНСКИЙ РАЙОН"},
        {"id": 249006, "name": "АЙТЕКЕБИЙСКИЙ РАЙОН"},
        {"id": 249152, "name": "БАЙГАНИНСКИЙ РАЙОН"},
        {"id": 249398, "name": "КАРГАЛИНСКИЙ РАЙОН"},
        {"id": 249439, "name": "ХОБДИНСКИЙ РАЙОН"},
        {"id": 249572, "name": "МАРТУКСКИЙ РАЙОН"},
        {"id": 249637, "name": "МУГАЛЖАРСКИЙ РАЙОН"},
        {"id": 249809, "name": "УИЛСКИЙ РАЙОН"},
        {"id": 249877, "name": "ТЕМИРСКИЙ РАЙОН"},
        {"id": 249978, "name": "ХРОМТАУСКИЙ РАЙОН"},
        {"id": 250101, "name": "ШАЛКАРСКИЙ РАЙОН"},
        {"id": 250353, "name": "ИРГИЗСКИЙ РАЙОН"},
    ],
    "Г.АСТАНА": [
        {"id": 268015,   "name": "РАЙОН АЛМАТЫ"},
        {"id": 268017,   "name": "РАЙОН ЕСИЛЬ"},
        {"id": 268019,   "name": "РАЙОН САРЫАРКА"},
        {"id": 20169322, "name": "РАЙОН БАЙҚОҢЫР"},
        {"id": 77216130, "name": "РАЙОН НҰРА"},
        {"id": 77356522, "name": "РАЙОН САРАЙШЫҚ"},
    ],
    "КАРАГАНДИНСКАЯ ОБЛАСТЬ": [
        {"id": 256620, "name": "КАРАГАНДА Г.А."},
        {"id": 256627, "name": "БАЛХАШ Г.А."},
        {"id": 256698, "name": "ПРИОЗЕРСК Г.А."},
        {"id": 256700, "name": "САРАНЬ Г.А."},
        {"id": 256709, "name": "ТЕМИРТАУ Г.А."},
        {"id": 256713, "name": "ШАХТИНСК Г.А."},
        {"id": 256721, "name": "АБАЙСКИЙ РАЙОН"},
        {"id": 256779, "name": "АКТОГАЙСКИЙ РАЙОН"},
        {"id": 257147, "name": "БУХАР-ЖЫРАУСКИЙ РАЙОН"},
        {"id": 257376, "name": "КАРКАРАЛИНСКИЙ РАЙОН"},
        {"id": 257880, "name": "НУРИНСКИЙ РАЙОН"},
        {"id": 257993, "name": "ОСАКАРОВСКИЙ РАЙОН"},
        {"id": 258445, "name": "ШЕТСКИЙ РАЙОН"},
    ],
    "Г.ШЫМКЕНТ": [
        {"id": 20242101, "name": "АБАЙСКИЙ РАЙОН"},
        {"id": 20242102, "name": "КАРАТАУСКИЙ РАЙОН"},
        {"id": 20242103, "name": "ЕНБЕКШИНСКИЙ РАЙОН"},
        {"id": 20242104, "name": "АЛЬ-ФАРАБИЙСКИЙ РАЙОН"},
        {"id": 77215757, "name": "РАЙОН ТУРАН"},
    ],
    "Г.АЛМАТЫ": [
        {"id": 268023,   "name": "АЛМАЛИНСКИЙ РАЙОН"},
        {"id": 268025,   "name": "АЛАТАУСКИЙ РАЙОН"},
        {"id": 268027,   "name": "АУЭЗОВСКИЙ РАЙОН"},
        {"id": 268029,   "name": "БОСТАНДЫКСКИЙ РАЙОН"},
        {"id": 268031,   "name": "ЖЕТЫСУСКИЙ РАЙОН"},
        {"id": 268033,   "name": "МЕДЕУСКИЙ РАЙОН"},
        {"id": 268035,   "name": "ТУРКСИБСКИЙ РАЙОН"},
        {"id": 17807940, "name": "НАУРЫЗБАЙСКИЙ РАЙОН"},
    ],
    "ПАВЛОДАРСКАЯ ОБЛАСТЬ": [
        {"id": 263010, "name": "ПАВЛОДАР Г.А."},
        {"id": 263024, "name": "АКСУ Г.А."},
        {"id": 263088, "name": "ЭКИБАСТУЗ Г.А."},
        {"id": 263165, "name": "АКТОГАЙСКИЙ РАЙОН"},
        {"id": 263218, "name": "БАЯНАУЛЬСКИЙ РАЙОН"},
        {"id": 263465, "name": "ЖЕЛЕЗИНСКИЙ РАЙОН"},
        {"id": 263531, "name": "ИРТЫШСКИЙ РАЙОН"},
        {"id": 263588, "name": "РАЙОН ТЕРЕҢКӨЛ"},
        {"id": 263659, "name": "РАЙОН АҚҚУЛЫ"},
        {"id": 263726, "name": "МАЙСКИЙ РАЙОН"},
        {"id": 263842, "name": "ПАВЛОДАРСКИЙ РАЙОН"},
        {"id": 263925, "name": "УСПЕНСКИЙ РАЙОН"},
        {"id": 263964, "name": "ЩЕРБАКТИНСКИЙ РАЙОН"},
    ],
    "ТУРКЕСТАНСКАЯ ОБЛАСТЬ": [
        {"id": 20243033, "name": "ТУРКЕСТАН Г.А."},
        {"id": 20243105, "name": "АРЫСЬ Г.А."},
        {"id": 20243145, "name": "КЕНТАУ Г.А."},
        {"id": 20243158, "name": "РАЙОН БАЙДИБЕКА"},
        {"id": 20243249, "name": "КАЗЫГУРТСКИЙ РАЙОН"},
        {"id": 20243457, "name": "МАКТААРАЛЬСКИЙ РАЙОН"},
        {"id": 20243664, "name": "ОРДАБАСЫНСКИЙ РАЙОН"},
        {"id": 20243732, "name": "ОТРАРСКИЙ РАЙОН"},
        {"id": 20243790, "name": "САЙРАМСКИЙ РАЙОН"},
        {"id": 20243860, "name": "САРЫАГАШСКИЙ РАЙОН"},
        {"id": 20244040, "name": "СУЗАКСКИЙ РАЙОН"},
        {"id": 20244091, "name": "ТОЛЕБИЙСКИЙ РАЙОН"},
        {"id": 20244159, "name": "ТЮЛЬКУБАССКИЙ РАЙОН"},
        {"id": 20244243, "name": "ШАРДАРИНСКИЙ РАЙОН"},
        {"id": 20256655, "name": "ЖЕТЫСАЙСКИЙ РАЙОН"},
        {"id": 20256656, "name": "КЕЛЕССКИЙ РАЙОН"},
        {"id": 77036054, "name": "РАЙОН САУРАН"},
    ],
    "ОБЛАСТЬ ҰЛЫТАУ": [
        {"id": 256636, "name": "ЖЕЗКАЗГАН Г.А."},
        {"id": 256694, "name": "КАРАЖАЛ Г.А."},
        {"id": 256704, "name": "САТПАЕВ Г.А."},
        {"id": 257257, "name": "ЖАНААРКИНСКИЙ РАЙОН"},
        {"id": 258074, "name": "УЛЫТАУСКИЙ РАЙОН"},
    ],
    "ВОСТОЧНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": [
        {"id": 264991, "name": "УСТЬ-КАМЕНОГОРСК Г.А."},
        {"id": 265011, "name": "РИДДЕР Г.А."},
        {"id": 266103, "name": "ГЛУБОКОВСКИЙ РАЙОН"},
        {"id": 266477, "name": "ЗАЙСАНСКИЙ РАЙОН"},
        {"id": 266711, "name": "РАЙОН АЛТАЙ"},
        {"id": 266872, "name": "КУРЧУМСКИЙ РАЙОН"},
        {"id": 267072, "name": "КАТОН-КАРАГАЙСКИЙ РАЙОН"},
        {"id": 267140, "name": "ТАРБАГАТАЙСКИЙ РАЙОН"},
        {"id": 267752, "name": "УЛАНСКИЙ РАЙОН"},
        {"id": 267955, "name": "ШЕМОНАИХИНСКИЙ РАЙОН"},
        {"id": 77208143, "name": "РАЙОН САМАР"},
        {"id": 77258303, "name": "РАЙОН МАРҚАКӨЛ"},
        {"id": 77258304, "name": "РАЙОН ҮЛКЕН НАРЫН"},
    ],
    "ЗАПАДНО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": [
        {"id": 253161, "name": "УРАЛЬСК Г.А."},
        {"id": 253176, "name": "АКЖАИКСКИЙ РАЙОН"},
        {"id": 253495, "name": "БУРЛИНСКИЙ РАЙОН"},
        {"id": 253546, "name": "ЖАНГАЛИНСКИЙ РАЙОН"},
        {"id": 253824, "name": "ЖАНИБЕКСКИЙ РАЙОН"},
        {"id": 253996, "name": "РАЙОН БӘЙТЕРЕК"},
        {"id": 254117, "name": "КАЗТАЛОВСКИЙ РАЙОН"},
        {"id": 254548, "name": "КАРАТОБИНСКИЙ РАЙОН"},
        {"id": 254695, "name": "БОКЕЙОРДИНСКИЙ РАЙОН"},
        {"id": 255132, "name": "СЫРЫМСКИЙ РАЙОН"},
        {"id": 255288, "name": "ТАСКАЛИНСКИЙ РАЙОН"},
        {"id": 255407, "name": "ТЕРЕКТИНСКИЙ РАЙОН"},
        {"id": 255505, "name": "ЧИНГИРЛАУСКИЙ РАЙОН"},
    ],
    "СЕВЕРО-КАЗАХСТАНСКАЯ ОБЛАСТЬ": [
        {"id": 264024, "name": "ПЕТРОПАВЛОВСК Г.А."},
        {"id": 264031, "name": "АЙЫРТАУСКИЙ РАЙОН"},
        {"id": 264140, "name": "АКЖАРСКИЙ РАЙОН"},
        {"id": 264178, "name": "РАЙОН МАГЖАНА ЖУМАБАЕВА"},
        {"id": 264274, "name": "ЕСИЛЬСКИЙ РАЙОН"},
        {"id": 264351, "name": "ЖАМБЫЛСКИЙ РАЙОН"},
        {"id": 264430, "name": "КЫЗЫЛЖАРСКИЙ РАЙОН"},
        {"id": 264529, "name": "МАМЛЮТСКИЙ РАЙОН"},
        {"id": 264587, "name": "РАЙОН ШАЛ АКЫНА"},
        {"id": 264646, "name": "АККАЙЫНСКИЙ РАЙОН"},
        {"id": 264693, "name": "ТАЙЫНШИНСКИЙ РАЙОН"},
        {"id": 264811, "name": "ТИМИРЯЗЕВСКИЙ РАЙОН"},
        {"id": 264854, "name": "УАЛИХАНОВСКИЙ РАЙОН"},
        {"id": 264901, "name": "РАЙОН ИМ.ГАБИТА МУСРЕПОВА"},
    ],
    "ОБЛАСТЬ АБАЙ": [
        {"id": 265009,   "name": "КУРЧАТОВ Г.А."},
        {"id": 265067,   "name": "СЕМЕЙ Г.А."},
        {"id": 265327,   "name": "АБАЙСКИЙ РАЙОН"},
        {"id": 265691,   "name": "АЯГОЗСКИЙ РАЙОН"},
        {"id": 265933,   "name": "БЕСКАРАГАЙСКИЙ РАЙОН"},
        {"id": 266012,   "name": "БОРОДУЛИХИНСКИЙ РАЙОН"},
        {"id": 266177,   "name": "ЖАРМИНСКИЙ РАЙОН"},
        {"id": 266788,   "name": "КОКПЕКТИНСКИЙ РАЙОН"},
        {"id": 267857,   "name": "УРДЖАРСКИЙ РАЙОН"},
        {"id": 77208142, "name": "РАЙОН АҚСУАТ"},
        {"id": 77258397, "name": "РАЙОН МАҚАНШЫ"},
        {"id": 77265110, "name": "РАЙОН ЖАҢАСЕМЕЙ"},
    ],
    "МАНГИСТАУСКАЯ ОБЛАСТЬ": [
        {"id": 260908, "name": "АКТАУ Г.А."},
        {"id": 260912, "name": "ЖАНАОЗЕН Г.А."},
        {"id": 260918, "name": "БЕЙНЕУСКИЙ РАЙОН"},
        {"id": 261032, "name": "КАРАКИЯНСКИЙ РАЙОН"},
        {"id": 261124, "name": "МАНГИСТАУСКИЙ РАЙОН"},
        {"id": 261352, "name": "МУНАЙЛИНСКИЙ РАЙОН"},
        {"id": 261369, "name": "ТУПКАРАГАНСКИЙ РАЙОН"},
    ],
    "АТЫРАУСКАЯ ОБЛАСТЬ": [
        {"id": 252312, "name": "АТЫРАУ Г.А."},
        {"id": 252371, "name": "ЖЫЛЫОЙСКИЙ РАЙОН"},
        {"id": 252445, "name": "ИНДЕРСКИЙ РАЙОН"},
        {"id": 252685, "name": "ИСАТАЙСКИЙ РАЙОН"},
        {"id": 252748, "name": "КУРМАНГАЗИНСКИЙ РАЙОН"},
        {"id": 252826, "name": "КЗЫЛКОГИНСКИЙ РАЙОН"},
        {"id": 253049, "name": "МАКАТСКИЙ РАЙОН"},
        {"id": 253062, "name": "МАХАМБЕТСКИЙ РАЙОН"},
    ],
    "КОСТАНАЙСКАЯ ОБЛАСТЬ": [
        {"id": 258743, "name": "КОСТАНАЙ Г.А."},
        {"id": 258745, "name": "АРКАЛЫК Г.А."},
        {"id": 258875, "name": "ЛИСАКОВСК Г.А."},
        {"id": 258881, "name": "РУДНЫЙ Г.А."},
        {"id": 258888, "name": "АЛТЫНСАРИНСКИЙ РАЙОН"},
        {"id": 258937, "name": "АМАНГЕЛЬДИНСКИЙ РАЙОН"},
        {"id": 259056, "name": "АУЛИЕКОЛЬСКИЙ РАЙОН"},
        {"id": 259122, "name": "ДЕНИСОВСКИЙ РАЙОН"},
        {"id": 259184, "name": "ЖАНГЕЛЬДИНСКИЙ РАЙОН"},
        {"id": 259303, "name": "ЖИТИКАРИНСКИЙ РАЙОН"},
        {"id": 259350, "name": "КАМЫСТИНСКИЙ РАЙОН"},
        {"id": 259391, "name": "КАРАБАЛЫКСКИЙ РАЙОН"},
        {"id": 259472, "name": "КАРАСУСКИЙ РАЙОН"},
        {"id": 259550, "name": "КОСТАНАЙСКИЙ РАЙОН"},
        {"id": 259642, "name": "МЕНДЫКАРИНСКИЙ РАЙОН"},
        {"id": 259726, "name": "НАУРЗУМСКИЙ РАЙОН"},
        {"id": 259813, "name": "САРЫКОЛЬСКИЙ РАЙОН"},
        {"id": 259873, "name": "РАЙОН БЕИМБЕТА МАЙЛИНА"},
        {"id": 259939, "name": "УЗУНКОЛЬСКИЙ РАЙОН"},
        {"id": 260015, "name": "ФЕДОРОВСКИЙ РАЙОН"},
    ],
    "КЫЗЫЛОРДИНСКАЯ ОБЛАСТЬ": [
        {"id": 260100, "name": "КЫЗЫЛОРДА Г.А."},
        {"id": 260132, "name": "БАЙКОНЫР Г.А."},
        {"id": 260134, "name": "АРАЛЬСКИЙ РАЙОН"},
        {"id": 260423, "name": "ЖАЛАГАШСКИЙ РАЙОН"},
        {"id": 260475, "name": "ЖАНАКОРГАНСКИЙ РАЙОН"},
        {"id": 260547, "name": "КАЗАЛИНСКИЙ РАЙОН"},
        {"id": 260631, "name": "КАРМАКШИНСКИЙ РАЙОН"},
        {"id": 260702, "name": "СЫРДАРЬИНСКИЙ РАЙОН"},
        {"id": 260770, "name": "ШИЕЛИЙСКИЙ РАЙОН"},
    ],
    "ОБЛАСТЬ ЖЕТІСУ": [
        {"id": 250503, "name": "ТАЛДЫКОРГАН Г.А."},
        {"id": 250538, "name": "ТЕКЕЛИ Г.А."},
        {"id": 250542, "name": "АКСУСКИЙ РАЙОН"},
        {"id": 250611, "name": "АЛАКОЛЬСКИЙ РАЙОН"},
        {"id": 251325, "name": "КЕРБУЛАКСКИЙ РАЙОН"},
        {"id": 251435, "name": "КОКСУСКИЙ РАЙОН"},
        {"id": 251488, "name": "КАРАТАЛЬСКИЙ РАЙОН"},
        {"id": 251666, "name": "ПАНФИЛОВСКИЙ РАЙОН"},
        {"id": 251880, "name": "САРКАНСКИЙ РАЙОН"},
        {"id": 252088, "name": "ЕСКЕЛЬДИНСКИЙ РАЙОН"},
    ],
    "ЖАМБЫЛСКАЯ ОБЛАСТЬ": [
        {"id": 255578, "name": "ТАРАЗ Г.А."},
        {"id": 255580, "name": "БАЙЗАКСКИЙ РАЙОН"},
        {"id": 255668, "name": "ЖАМБЫЛСКИЙ РАЙОН"},
        {"id": 255780, "name": "ЖУАЛЫНСКИЙ РАЙОН"},
        {"id": 255879, "name": "КОРДАЙСКИЙ РАЙОН"},
        {"id": 255978, "name": "РАЙОН ТУРАРА РЫСКУЛОВА"},
        {"id": 256156, "name": "МЕРКЕНСКИЙ РАЙОН"},
        {"id": 256255, "name": "МОЙЫНКУМСКИЙ РАЙОН"},
        {"id": 256365, "name": "САРЫСУСКИЙ РАЙОН"},
        {"id": 256442, "name": "ТАЛАССКИЙ РАЙОН"},
        {"id": 256504, "name": "ШУСКИЙ РАЙОН"},
    ],
}

# ─────────────────────────────────────────────────────────────────────
# ОСНОВНОЙ ЦИКЛ
# ─────────────────────────────────────────────────────────────────────

def main():
    log("=" * 60)
    log("ПАРСИНГ МЕСЯЧНЫХ ДАННЫХ — ВСЕ РЕГИОНЫ КАЗАХСТАНА")
    log("=" * 60)

    # Загружаем прогресс предыдущего запуска
    done_districts = set()
    if os.path.exists(DONE_FILE):
        with open(DONE_FILE, "r", encoding="utf-8") as f:
            done_districts = set(line.strip() for line in f if line.strip())
        log(f"▶ Продолжаем. Уже готово: {len(done_districts)} районов")

    grand_total = []
    if os.path.exists(PROGRESS_FILE) and done_districts:
        grand_total = pd.read_csv(PROGRESS_FILE, encoding="utf-8-sig").to_dict("records")
        log(f"  Загружено {len(grand_total)} строк из предыдущего запуска")

    if not check_session():
        log("СТОП: сначала обнови Cookie!")
        raise SystemExit(1)

    try:
        total_oblasts = len(ALL_REGIONS)
        for oblast_num, (oblast_name, districts) in enumerate(ALL_REGIONS.items(), 1):
            log(f"\n{'='*60}")
            log(f"[{oblast_num}/{total_oblasts}] {oblast_name} ({len(districts)} районов)")
            log(f"{'='*60}")

            oblast_data = []

            for i, item in enumerate(districts, 1):
                district_key = f"{oblast_name}|{item['name']}"

                if district_key in done_districts:
                    log(f"  [{i}/{len(districts)}] {item['name']} — пропущен (уже есть)")
                    continue

                log(f"  [{i}/{len(districts)}] {item['name']}...")

                p_terms = f"{item['id']},{TERMS_SUFFIX}"
                rows = fetch_recursive(p_terms, parent_id="")

                for row in rows:
                    row["oblast"] = oblast_name
                    row["region"] = item["name"]

                oblast_data.extend(rows)
                grand_total.extend(rows)

                log(f"    ✓ {len(rows)} строк")

                done_districts.add(district_key)
                with open(DONE_FILE, "a", encoding="utf-8") as f:
                    f.write(district_key + "\n")

                # Сохраняем чекпоинт после каждого района
                pd.DataFrame(grand_total).to_csv(PROGRESS_FILE, index=False, encoding="utf-8-sig")

            # Сохраняем файл по области
            if oblast_data:
                safe_name = (
                    oblast_name
                    .replace(" ", "_").replace(".", "")
                    .replace("Ұ", "U").replace("É", "ZH").replace("Ж", "ZH")
                )
                oblast_path = os.path.join(OUTPUT_DIR, f"{safe_name}.csv")

                if os.path.exists(oblast_path):
                    existing = pd.read_csv(oblast_path, encoding="utf-8-sig")
                    combined = pd.concat([existing, pd.DataFrame(oblast_data)], ignore_index=True)
                    combined.to_csv(oblast_path, index=False, encoding="utf-8-sig")
                else:
                    pd.DataFrame(oblast_data).to_csv(oblast_path, index=False, encoding="utf-8-sig")

                log(f"  💾 {oblast_path} ({len(oblast_data)} строк)")

    except CookieExpiredError as e:
        log(f"\n{'!'*60}")
        log(f"  🛑 СТОП: куки устарели в процессе работы!")
        log(f"  Причина: {e}")
        log(f"  Прогресс сохранён. Обнови Cookie и запусти снова.")
        log(f"  Продолжит с места остановки (см. done.txt).")
        log(f"{'!'*60}")
        if grand_total:
            pd.DataFrame(grand_total).to_csv(PROGRESS_FILE, index=False, encoding="utf-8-sig")
        raise SystemExit(1)

    # ── Финальный файл ────────────────────────────────────────────────
    final_df = pd.DataFrame(grand_total)

    # Оставляем period-колонки + служебные
    base_cols   = ["oblast", "region", "depth", "parent_text", "text", "leaf", "id"]
    period_cols = list(TARGET_PERIODS.keys())
    all_extra   = [c for c in final_df.columns if c not in base_cols and c not in period_cols]

    keep_cols = (
        [c for c in base_cols if c in final_df.columns] +
        [c for c in period_cols if c in final_df.columns] +
        all_extra
    )
    final_df = final_df[keep_cols]
    final_df.to_csv(FINAL_CSV, index=False, encoding="utf-8-sig")

    log(f"\n{'='*60}")
    log(f"  ✅ ГОТОВО!")
    log(f"  Областей:  {final_df['oblast'].nunique() if not final_df.empty else 0}")
    log(f"  Районов:   {final_df['region'].nunique() if not final_df.empty else 0}")
    log(f"  Строк:     {len(final_df)}")
    log(f"  Запросов:  {stats['total_requests']}")
    log(f"  Пустых:    {stats['empty_responses']}")
    log(f"  Файл:      {FINAL_CSV}")
    log(f"{'='*60}")

    log("\nПокрытие по областям:")
    coverage = final_df.groupby("oblast")["region"].nunique().sort_index()
    for oblast, cnt in coverage.items():
        log(f"  {oblast}: {cnt} районов")


if __name__ == "__main__":
    main()
