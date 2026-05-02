"""
handlers/menu.py — Главное меню и навигация бота.
"""

import logging

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from core.storage import load_master, load_state, get_cookie, get_full_period_map
from core.periods import fmt_period, period_sort_key, is_yearly_key, is_monthly_key

logger = logging.getLogger(__name__)

# ── Клавиатуры ────────────────────────────────────────────────────────────────

TOP_MENU_KB = [
    [InlineKeyboardButton("📊 ИОК — Инвестиции в основной капитал", callback_data="iok_menu")],
    [InlineKeyboardButton("📈 ИПЦ — Инфляция (цены)",               callback_data="cpi_menu")],
]

IOK_MENU_KB = [
    [InlineKeyboardButton("🔍 Проверить наличие новых данных",  callback_data="check")],
    [InlineKeyboardButton("⬇️ Загрузить новые данные",           callback_data="update_auto")],
    [InlineKeyboardButton("📥 Получить файл Excel",              callback_data="download")],
    [InlineKeyboardButton("⚙️ Дополнительно",                   callback_data="iok_more")],
    [InlineKeyboardButton("← Назад",                            callback_data="main_menu")],
]

IOK_MORE_KB = [
    [InlineKeyboardButton("📅 Загрузить за конкретный месяц",   callback_data="manual_start")],
    [InlineKeyboardButton("📆 Загрузить годовые данные",         callback_data="yearly_start")],
    [InlineKeyboardButton("🔄 Продолжить прерванную загрузку",  callback_data="resume")],
    [InlineKeyboardButton("📊 Состояние базы данных",           callback_data="status")],
    [InlineKeyboardButton("🔑 Обновить Cookie",                 callback_data="set_cookie")],
    [InlineKeyboardButton("← Назад",                           callback_data="iok_menu")],
]


# ── Вспомогательные функции ───────────────────────────────────────────────────

def _cookie_status() -> str:
    return "✅ установлен" if get_cookie() else "❌ не установлен — обнови через ⚙️ Дополнительно"


def _iok_status_lines() -> str:
    master   = load_master()
    state    = load_state()
    loaded_m = state.get("loaded_periods", [])
    loaded_y = state.get("loaded_yearly",  [])

    if master.empty:
        return "📦 База данных пустая — нажми *Загрузить новые данные*"

    lines = [f"📦 Записей в базе: *{len(master):,}*"]

    all_loaded = sorted(loaded_m + loaded_y, key=period_sort_key)
    if all_loaded:
        lines.append(f"📅 Период: {fmt_period(all_loaded[0])} — {fmt_period(all_loaded[-1])}")

    last_upd = state.get("last_update")
    if last_upd:
        lines.append(f"🕐 Обновлено: {last_upd[:16].replace('T', ' ')}")

    return "\n".join(lines)


def _build_status_text(master, state: dict) -> str:
    from core.periods import is_period_key
    lines = ["*📊 Состояние базы данных — ИОК*\n"]

    if master.empty:
        lines.append("📦 База данных пустая\n\nЧтобы начать — нажми *Загрузить новые данные* в меню ИОК.")
    else:
        lines.append(f"📦 Записей в базе: *{len(master):,}*\n")
        all_cols   = [c for c in master.columns if is_period_key(c)]
        year_cols  = sorted([c for c in all_cols if is_yearly_key(c)],  key=period_sort_key)
        month_cols = sorted([c for c in all_cols if is_monthly_key(c)], key=period_sort_key)
        if year_cols:
            lines.append(f"📆 Годовые данные: {fmt_period(year_cols[0])} — {fmt_period(year_cols[-1])}")
        if month_cols:
            lines.append(f"📅 Месячные данные: {fmt_period(month_cols[0])} — {fmt_period(month_cols[-1])}")

    last_upd = state.get("last_update")
    if last_upd:
        lines.append(f"\n🕐 Последнее обновление: {last_upd[:16].replace('T', ' ')}")
    last_chk = state.get("last_check")
    if last_chk:
        lines.append(f"🔍 Последняя проверка: {last_chk[:16].replace('T', ' ')}")

    lines.append(f"\n🔑 Cookie: {_cookie_status()}")
    return "\n".join(lines)


# ── Команды ───────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    cookie_hint = "" if get_cookie() else "\n\n⚠️ Cookie не установлен — зайди в ИОК → Дополнительно → Обновить Cookie"
    await update.message.reply_text(
        f"👋 *Добро пожаловать!*\n\n"
        f"Этот бот скачивает статистику с сайта *taldau.stat.gov.kz* "
        f"и формирует удобные Excel-файлы.\n\n"
        f"Выберите раздел:{cookie_hint}",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(TOP_MENU_KB),
    )


async def cmd_help(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "📖 *Справка*\n\n"
        "*Что делает бот:*\n"
        "Бот заходит на taldau.stat.gov.kz и скачивает данные в формате Excel.\n\n"
        "*Разделы:*\n"
        "📊 *ИОК* — Инвестиции в основной капитал (по регионам и месяцам)\n"
        "📈 *ИПЦ* — Индексы потребительских цен (инфляция по 21 региону)\n\n"
        "*Зачем нужен Cookie?*\n"
        "Сайт taldau.stat.gov.kz требует авторизации. Cookie — это ключ доступа, "
        "который нужно обновлять примерно раз в несколько недель.\n\n"
        "*Как получить Cookie:*\n"
        "1. Открой taldau.stat.gov.kz в браузере и войди в аккаунт\n"
        "2. Нажми F12 → вкладка *Network*\n"
        "3. Обнови страницу, найди запрос `GetIndexTreeData`\n"
        "4. В разделе *Headers* скопируй строку после `Cookie:`\n"
        "5. Вставь в бот через меню *Дополнительно → Обновить Cookie*\n\n"
        "*Команды:*\n"
        "/start — главное меню\n"
        "/status — состояние базы данных\n"
        "/stop — остановить загрузку\n"
        "/help — эта справка",
        parse_mode="Markdown",
    )


async def cmd_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        _build_status_text(load_master(), load_state()),
        parse_mode="Markdown",
    )


# ── Callbacks ─────────────────────────────────────────────────────────────────

async def cb_main_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    await q.edit_message_text(
        "Выберите раздел:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(TOP_MENU_KB),
    )


async def cb_iok_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    await q.edit_message_text(
        f"*📊 ИОК — Инвестиции в основной капитал*\n\n"
        f"{_iok_status_lines()}\n\n"
        f"🔑 Cookie: {_cookie_status()}",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(IOK_MENU_KB),
    )


async def cb_iok_more(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    await q.edit_message_text(
        "*⚙️ Дополнительные возможности — ИОК*\n\n"
        "Здесь можно загрузить данные за конкретный период, "
        "продолжить прерванную загрузку или обновить Cookie.",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(IOK_MORE_KB),
    )


async def cb_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()

    from config.settings import OUTPUT_DIR
    text = _build_status_text(load_master(), load_state())
    xlsx = sorted(OUTPUT_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    if xlsx:
        text += f"\n\n📁 Файлов Excel: {len(xlsx)}\nПоследний: `{xlsx[0].name}`"

    await q.edit_message_text(
        text,
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("◀️ Назад", callback_data="iok_more")],
        ]),
    )
