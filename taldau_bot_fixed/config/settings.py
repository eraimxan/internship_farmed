"""
config/settings.py — Централизованные настройки бота
Все константы, пути и параметры API.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

# ─────────────────────────────────────────────
# TELEGRAM
# ─────────────────────────────────────────────
BOT_TOKEN: str = os.getenv("BOT_TOKEN", "6395810050:AAFahhL0kdFQwlUYd3k213TJalGy_xhyROE")

# ─────────────────────────────────────────────
# ПУТИ К ФАЙЛАМ
# ─────────────────────────────────────────────
BASE_DIR    = Path(__file__).resolve().parent.parent
DATA_DIR    = BASE_DIR / "data"
OUTPUT_DIR  = BASE_DIR / "output"
MASTER_CSV  = DATA_DIR / "master.csv"
STATE_FILE  = DATA_DIR / "state.json"
COOKIE_FILE = DATA_DIR / "cookie.txt"
LOG_FILE    = BASE_DIR / "bot.log"

DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────
# API TALDAU
# ─────────────────────────────────────────────
BASE_URL     = "https://taldau.stat.gov.kz/ru/NewIndex/GetIndexTreeData"
INDEX_ID     = "701830"
TERM_ID      = "741885"

# Типы периодов на сайте taldau
PERIOD_ID_MONTHLY = "8"    # Месяц с накоплением (Янв, Янв-Фев, ...)
PERIOD_ID_YEARLY  = "7"    # Годовые данные (2019, 2020, ...)

# Для обратной совместимости
PERIOD_ID = PERIOD_ID_MONTHLY

# ── Параметры для МЕСЯЧНЫХ данных (period_id=8) ──
MONTHLY_DIC_IDS      = "68,61,3028"
MONTHLY_TERMS_SUFFIX = "741926,741885"
MONTHLY_IDX          = "2"

# ── Параметры для ГОДОВЫХ данных (period_id=7) ──
YEARLY_DIC_IDS      = "68,776,59,3028,4043"
YEARLY_TERMS_SUFFIX = "741917,741907,741885,19202525"
YEARLY_IDX          = "3"

# Обратная совместимость
DIC_IDS      = MONTHLY_DIC_IDS
TERMS_SUFFIX = MONTHLY_TERMS_SUFFIX

# ─────────────────────────────────────────────
# ПАРАМЕТРЫ ПАРСИНГА
# ─────────────────────────────────────────────
REQUEST_DELAY          = 0.3  # секунды между запросами
RETRY_COUNT            = 4    # попыток при ошибке
REQUEST_TIMEOUT        = 90   # timeout одного запроса (секунды)
MAX_DEPTH              = 8    # макс глубина рекурсии
MAX_WORKERS            = 3    # параллельных воркеров парсинга
PROGRESS_SAVE_INTERVAL = 5    # сохранять прогресс каждые N районов

# Максимальный размер файла для Telegram Bot API (байты)
TELEGRAM_FILE_LIMIT = 49 * 1024 * 1024  # 49 МБ

# ─────────────────────────────────────────────
# КЛЮЧЕВЫЕ ПОЛЯ (идентификаторы строки)
# ─────────────────────────────────────────────
KEY_COLS = ["oblast", "region", "depth", "text"]

# ─────────────────────────────────────────────
# НАЗВАНИЯ МЕСЯЦЕВ
# ─────────────────────────────────────────────
MONTH_RU = {
    1: "Янв", 2: "Фев",  3: "Мар", 4: "Апр",
    5: "Май", 6: "Июн",  7: "Июл", 8: "Авг",
    9: "Сен", 10: "Окт", 11: "Ноя", 12: "Дек",
}

# ─────────────────────────────────────────────
# EXCEL СТИЛИ
# ─────────────────────────────────────────────
LEVEL_COLORS = {
    0: ("1F3864", "FFFFFF", True,  11),
    1: ("2E75B6", "FFFFFF", True,  10),
    2: ("BDD7EE", "1F3864", True,  9),
    3: ("DDEBF7", "1F3864", False, 9),
    4: ("EEF4FB", "2E3A4E", False, 9),
    5: ("F5F9FD", "2E3A4E", False, 9),
    6: ("FFFFFF", "2E3A4E", False, 9),
}
