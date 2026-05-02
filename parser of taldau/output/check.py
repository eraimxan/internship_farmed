"""
verify_and_patch.py
───────────────────────────────────────────────────────────────────────────────
Сверяет LOCAL датасет (ALL_KAZAKHSTAN_FINAL.csv) с данными с сайта
taldau.stat.gov.kz и исправляет/дополняет расхождения.

Логика:
  1. Для каждого (oblast, region) → строит p_terms и запрашивает API
  2. Сравнивает каждую строку по id + все числовые колонки
  3. Если значение отличается или было NaN → перезаписывает
  4. Сверяет текстовые поля (text, measureName, parent_text) и исправляет
  5. Если строки нет в локальном датасете → добавляет
  6. Если строка есть локально, но отсутствует на сайте → удаляет
  7. Сохраняет патченный файл + детальный лог расхождений

Запуск:
  python verify_and_patch.py

⚠️  Перед запуском обнови Cookie в разделе КОНФИГУРАЦИЯ!
"""

import requests
import pandas as pd
import time
import os
import json
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────
# КОНФИГУРАЦИЯ — обновляй Cookie перед каждым запуском!
# ─────────────────────────────────────────────────────────────────────

COOKIE = "_ym_uid=1758390801129425250; _ym_d=1758390801; ASP.NET_SessionId=ld5crs1w31h5z1lj3lskkiaw"

INPUT_CSV   = "ALL_KAZAKHSTAN_FINAL.csv"   # исходный файл
OUTPUT_CSV  = "ALL_KAZAKHSTAN_PATCHED.csv" # результат
DIFF_CSV    = "diff_report.csv"            # лог всех расхождений
LOG_FILE    = "verify_patch.log"
DONE_FILE   = "verify_done.txt"            # уже проверенные регионы (для продолжения)

# Задержка между запросами (сек). Увеличь если сайт начинает блокировать.
DELAY = 1.2

# Максимальная глубина рекурсии при обходе дерева
MAX_DEPTH = 10

# ─────────────────────────────────────────────────────────────────────
# API-параметры (из оригинального скрипта)
# ─────────────────────────────────────────────────────────────────────

BASE         = "https://taldau.stat.gov.kz/ru/NewIndex/GetIndexTreeData"
INDEX_ID     = "701830"
PERIOD_ID    = "7"
DIC_IDS      = "68,776,59,3028,4043"
TERM_ID      = "741885"
TERMS_SUFFIX = "741917,741907,741885,19202525"

# ─────────────────────────────────────────────────────────────────────
# Числовые колонки (год-данные)
# ─────────────────────────────────────────────────────────────────────

DATA_COLS = ['y122020','122020','y122021','122021','y122019','122019',
             'y122024','122024','y122023','122023','y122022','122022']

# Текстовые колонки для сверки
TEXT_COLS = ['text', 'measureName', 'parent_text']

# ─────────────────────────────────────────────────────────────────────
# Сессия
# ─────────────────────────────────────────────────────────────────────

session = requests.Session()
session.headers.update({
    "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/123.0.0.0 Safari/537.36",
    "Accept":          "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    "X-Requested-With":"XMLHttpRequest",
    "Referer":         "https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/701830",
    "Cookie":          COOKIE,
})

# ─────────────────────────────────────────────────────────────────────
# Логгер
# ─────────────────────────────────────────────────────────────────────

def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

# ─────────────────────────────────────────────────────────────────────
# Очистка значений (как в оригинале)
# ─────────────────────────────────────────────────────────────────────

def parse_val(v):
    """Преобразует строковое значение из API в число или None."""
    if v is None or v == "":
        return None
    try:
        return int(v)
    except (ValueError, TypeError):
        try:
            return float(v)
        except (ValueError, TypeError):
            return v

def clean_row(row: dict) -> dict:
    cleaned = {}
    for k, v in row.items():
        is_data_col = (
            (k.startswith("y") and k[1:].isdigit()) or
            k.isdigit() or
            (len(k) >= 2 and k[:2].isdigit())
        )
        cleaned[k] = parse_val(v) if is_data_col else v
    return cleaned

# ─────────────────────────────────────────────────────────────────────
# Запрос к API
# ─────────────────────────────────────────────────────────────────────

class CookieExpiredError(Exception):
    pass

stats = {"requests": 0, "empty": 0, "errors": 0}

def fetch(p_terms, parent_id="", retries=5, is_root=False):
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
            time.sleep(DELAY)

            if r.status_code in (401, 403):
                raise CookieExpiredError(f"HTTP {r.status_code}")

            if r.status_code != 200:
                time.sleep(DELAY * (attempt + 1))
                continue

            text = r.text.strip()
            if len(text) <= 2 or text in ("[]", "null", ""):
                stats["empty"] += 1
                if is_root:
                    raise CookieExpiredError("Корневой запрос вернул пустоту — куки устарели!")
                if attempt < retries - 1:
                    time.sleep(DELAY * (attempt + 2))
                    continue
                return []

            result = r.json()
            return result if isinstance(result, list) else []

        except CookieExpiredError:
            raise
        except requests.exceptions.ReadTimeout:
            time.sleep(10 * (attempt + 1))
        except Exception as e:
            stats["errors"] += 1
            log(f"  ❌ Ошибка: {e}")
            time.sleep(DELAY * 2)
    return []


def fetch_recursive(p_terms, parent_id="", depth=0, parent_text=""):
    if depth > MAX_DEPTH:
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
            )
            all_rows.extend(children)
    return all_rows

# ─────────────────────────────────────────────────────────────────────
# Проверка сессии
# ─────────────────────────────────────────────────────────────────────

def check_session():
    log("Проверяем сессию...")
    try:
        rows = fetch(f"247784,{TERMS_SUFFIX}", parent_id="", is_root=True)
    except CookieExpiredError:
        log("  ❌ Куки устарели!")
        return False
    if rows:
        log(f"  ✅ Сессия работает! Строк в тесте: {len(rows)}")
        return True
    log("  ❌ 0 строк — куки устарели!")
    return False

# ─────────────────────────────────────────────────────────────────────
# Сравнение значений
# ─────────────────────────────────────────────────────────────────────

def vals_equal(local_val, remote_val):
    """
    True если значения считаются одинаковыми.
    Оба None → равны. Один None → НЕ равны (нужно патчить).
    Числа сравниваем через float.
    """
    # оба пусты
    if pd.isna(local_val) and remote_val is None:
        return True
    if pd.isna(local_val) or remote_val is None:
        return False
    try:
        return float(local_val) == float(remote_val)
    except (TypeError, ValueError):
        return str(local_val) == str(remote_val)

def texts_equal(local_val, remote_val):
    """Сравнение текстовых полей: NaN и None считаются одинаковыми."""
    local_empty  = local_val is None or (isinstance(local_val, float) and pd.isna(local_val)) or local_val == ""
    remote_empty = remote_val is None or remote_val == ""
    if local_empty and remote_empty:
        return True
    if local_empty or remote_empty:
        return False
    return str(local_val).strip() == str(remote_val).strip()

# ─────────────────────────────────────────────────────────────────────
# Регионы (скопировано из оригинала)
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
        {"id": 247784,   "name": "КОКШЕТАУ Г.А."},
        {"id": 247791,   "name": "СТЕПНОГОРСК Г.А."},
        {"id": 247807,   "name": "АККОЛЬСКИЙ РАЙОН"},
        {"id": 247855,   "name": "АРШАЛЫНСКИЙ РАЙОН"},
        {"id": 247904,   "name": "АСТРАХАНСКИЙ РАЙОН"},
        {"id": 247961,   "name": "АТБАСАРСКИЙ РАЙОН"},
        {"id": 248026,   "name": "БУЛАНДЫНСКИЙ РАЙОН"},
        {"id": 248076,   "name": "ЕГИНДЫКОЛЬСКИЙ РАЙОН"},
        {"id": 248100,   "name": "РАЙОН БИРЖАН САЛ"},
        {"id": 248165,   "name": "ЕРЕЙМЕНТАУСКИЙ РАЙОН"},
        {"id": 248261,   "name": "ЕСИЛЬСКИЙ РАЙОН"},
        {"id": 248322,   "name": "ЖАКСЫНСКИЙ РАЙОН"},
        {"id": 248378,   "name": "ЖАРКАИНСКИЙ РАЙОН"},
        {"id": 248433,   "name": "ЗЕРЕНДИНСКИЙ РАЙОН"},
        {"id": 248558,   "name": "КОРГАЛЖЫНСКИЙ РАЙОН"},
        {"id": 248592,   "name": "САНДЫКТАУСКИЙ РАЙОН"},
        {"id": 248655,   "name": "ЦЕЛИНОГРАДСКИЙ РАЙОН"},
        {"id": 248729,   "name": "ШОРТАНДИНСКИЙ РАЙОН"},
        {"id": 248775,   "name": "БУРАБАЙСКИЙ РАЙОН"},
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
        {"id": 264991,   "name": "УСТЬ-КАМЕНОГОРСК Г.А."},
        {"id": 265011,   "name": "РИДДЕР Г.А."},
        {"id": 266103,   "name": "ГЛУБОКОВСКИЙ РАЙОН"},
        {"id": 266477,   "name": "ЗАЙСАНСКИЙ РАЙОН"},
        {"id": 266711,   "name": "РАЙОН АЛТАЙ"},
        {"id": 266872,   "name": "КУРЧУМСКИЙ РАЙОН"},
        {"id": 267072,   "name": "КАТОН-КАРАГАЙСКИЙ РАЙОН"},
        {"id": 267140,   "name": "ТАРБАГАТАЙСКИЙ РАЙОН"},
        {"id": 267752,   "name": "УЛАНСКИЙ РАЙОН"},
        {"id": 267955,   "name": "ШЕМОНАИХИНСКИЙ РАЙОН"},
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
# ЗАГРУЗКА ЛОКАЛЬНОГО ДАТАСЕТА
# ─────────────────────────────────────────────────────────────────────

log(f"Загружаем {INPUT_CSV}...")
df_local = pd.read_csv(INPUT_CSV, encoding="utf-8-sig", low_memory=False)
# Приводим числовые колонки к float (NaN остаются NaN)
for c in DATA_COLS:
    if c in df_local.columns:
        df_local[c] = pd.to_numeric(df_local[c], errors="coerce")

log(f"  Строк: {len(df_local)}, колонок: {len(df_local.columns)}")

# ─────────────────────────────────────────────────────────────────────
# ПРОВЕРКА СЕССИИ
# ─────────────────────────────────────────────────────────────────────

if not check_session():
    log("СТОП: обнови Cookie в переменной COOKIE и запусти снова!")
    raise SystemExit(1)

# ─────────────────────────────────────────────────────────────────────
# ПРОДОЛЖЕНИЕ (если прерывалось)
# ─────────────────────────────────────────────────────────────────────

done_regions = set()
if os.path.exists(DONE_FILE):
    with open(DONE_FILE, "r", encoding="utf-8") as f:
        done_regions = set(ln.strip() for ln in f if ln.strip())
    log(f"Продолжаем: уже проверено {len(done_regions)} регионов")

# Работаем с копией датасета
df_patched = df_local.copy()

# Накапливаем diff-отчёт
diff_records = []
if os.path.exists(DIFF_CSV) and done_regions:
    diff_records = pd.read_csv(DIFF_CSV, encoding="utf-8-sig").to_dict("records")
    log(f"  Загружено {len(diff_records)} записей из предыдущего diff-отчёта")

# Счётчики
total_checked = 0
total_patched = 0
total_added   = 0
total_deleted = 0

# ─────────────────────────────────────────────────────────────────────
# ГЛАВНЫЙ ЦИКЛ
# ─────────────────────────────────────────────────────────────────────

try:
    n_oblasts = len(ALL_REGIONS)
    for obl_num, (oblast_name, districts) in enumerate(ALL_REGIONS.items(), 1):
        log(f"\n{'='*60}")
        log(f"[{obl_num}/{n_oblasts}] {oblast_name}  ({len(districts)} р-нов)")

        for d_num, item in enumerate(districts, 1):
            region_key = f"{oblast_name}|{item['name']}"

            if region_key in done_regions:
                log(f"  [{d_num}/{len(districts)}] {item['name']} — пропущен")
                continue

            log(f"  [{d_num}/{len(districts)}] {item['name']}...")

            # Получаем свежие данные с сайта
            p_terms = f"{item['id']},{TERMS_SUFFIX}"
            remote_rows = fetch_recursive(p_terms, parent_id="")

            if not remote_rows:
                log(f"    ⚠️ Сайт вернул 0 строк — пропускаем")
                done_regions.add(region_key)
                with open(DONE_FILE, "a", encoding="utf-8") as f:
                    f.write(region_key + "\n")
                continue

            # Индексируем локальные строки по id
            local_mask = (
                (df_patched["oblast"] == oblast_name) &
                (df_patched["region"] == item["name"])
            )
            local_subset = df_patched[local_mask]
            local_index  = {int(r["id"]): idx for idx, r in local_subset.iterrows()}

            # id-ы которые вернул сайт — для поиска лишних локальных строк
            remote_ids = set()

            region_patched = 0
            region_added   = 0
            region_deleted = 0

            for remote in remote_rows:
                rid = int(remote.get("id", -1))
                remote_ids.add(rid)
                total_checked += 1

                if rid in local_index:
                    # ── Строка есть — сравниваем все колонки ─────────
                    loc_idx = local_index[rid]
                    changed_cols = []

                    # Числовые колонки
                    for col in DATA_COLS:
                        if col not in df_patched.columns:
                            continue
                        local_val  = df_patched.at[loc_idx, col]
                        remote_val = remote.get(col)
                        if not vals_equal(local_val, remote_val):
                            changed_cols.append({
                                "col":        col,
                                "local_val":  local_val,
                                "remote_val": remote_val,
                            })
                            df_patched.at[loc_idx, col] = remote_val

                    # Текстовые колонки
                    for col in TEXT_COLS:
                        if col not in df_patched.columns:
                            continue
                        local_val  = df_patched.at[loc_idx, col]
                        remote_val = remote.get(col)
                        if not texts_equal(local_val, remote_val):
                            changed_cols.append({
                                "col":        col,
                                "local_val":  local_val,
                                "remote_val": remote_val,
                            })
                            df_patched.at[loc_idx, col] = remote_val

                    if changed_cols:
                        region_patched += 1
                        total_patched  += 1
                        for cc in changed_cols:
                            diff_records.append({
                                "oblast":      oblast_name,
                                "region":      item["name"],
                                "row_id":      rid,
                                "text":        remote.get("text", ""),
                                "col":         cc["col"],
                                "local_val":   cc["local_val"],
                                "remote_val":  cc["remote_val"],
                                "action":      "PATCHED",
                            })
                else:
                    # ── Строки нет — добавляем ────────────────────────
                    new_row = dict(remote)
                    new_row["oblast"]  = oblast_name
                    new_row["region"]  = item["name"]
                    for col in df_patched.columns:
                        if col not in new_row:
                            new_row[col] = None

                    df_patched = pd.concat(
                        [df_patched, pd.DataFrame([new_row])],
                        ignore_index=True
                    )
                    region_added += 1
                    total_added  += 1

                    diff_records.append({
                        "oblast":     oblast_name,
                        "region":     item["name"],
                        "row_id":     rid,
                        "text":       remote.get("text", ""),
                        "col":        "—",
                        "local_val":  "ОТСУТСТВУЕТ",
                        "remote_val": "—",
                        "action":     "ADDED",
                    })

            # ── Удаляем строки которых больше нет на сайте ───────────
            for local_id, loc_idx in local_index.items():
                if local_id not in remote_ids:
                    row_text = df_patched.at[loc_idx, "text"] if "text" in df_patched.columns else ""
                    diff_records.append({
                        "oblast":     oblast_name,
                        "region":     item["name"],
                        "row_id":     local_id,
                        "text":       row_text,
                        "col":        "—",
                        "local_val":  "УДАЛЕНО (нет на сайте)",
                        "remote_val": "—",
                        "action":     "DELETED",
                    })
                    df_patched = df_patched.drop(index=loc_idx)
                    region_deleted += 1
                    total_deleted  += 1

            df_patched = df_patched.reset_index(drop=True)

            log(f"    ✓ проверено {len(remote_rows)}, "
                f"исправлено: {region_patched}, "
                f"добавлено: {region_added}, "
                f"удалено: {region_deleted}")

            # Сохраняем прогресс
            done_regions.add(region_key)
            with open(DONE_FILE, "a", encoding="utf-8") as f:
                f.write(region_key + "\n")

            # Промежуточное сохранение каждые 10 регионов
            if len(done_regions) % 10 == 0:
                df_patched.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
                pd.DataFrame(diff_records).to_csv(DIFF_CSV, index=False, encoding="utf-8-sig")
                log(f"    💾 Промежуточное сохранение ({len(done_regions)} регионов)")

except CookieExpiredError as e:
    log(f"\n{'!'*60}")
    log(f"  🛑 СТОП: куки устарели!")
    log(f"  Причина: {e}")
    log(f"  Обнови Cookie и запусти снова — продолжит с места остановки.")
    log(f"{'!'*60}")
    # Сохраняем всё что успели
    df_patched.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    pd.DataFrame(diff_records).to_csv(DIFF_CSV, index=False, encoding="utf-8-sig")
    raise SystemExit(1)

# ─────────────────────────────────────────────────────────────────────
# ФИНАЛЬНОЕ СОХРАНЕНИЕ
# ─────────────────────────────────────────────────────────────────────

df_patched.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
pd.DataFrame(diff_records).to_csv(DIFF_CSV, index=False, encoding="utf-8-sig")

log(f"\n{'='*60}")
log(f" ГОТОВО!")
log(f"   Строк проверено:    {total_checked}")
log(f"   Строк исправлено:   {total_patched}")
log(f"   Строк добавлено:    {total_added}")
log(f"   Строк удалено:      {total_deleted}")
log(f"   API-запросов:       {stats['requests']}")
log(f"   Пустых ответов:     {stats['empty']}")
log(f"   Ошибок:             {stats['errors']}")
log(f"")
log(f"   Патченный файл:     {OUTPUT_CSV}")
log(f"   Лог расхождений:    {DIFF_CSV}")
log(f"   Лог выполнения:     {LOG_FILE}")
log(f"{'='*60}")