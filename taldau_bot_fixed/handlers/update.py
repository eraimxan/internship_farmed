"""
handlers/update.py — Автоматическое и ручное обновление данных ИОК.

Логика обновления:
  - Сравниваем периоды с сайта против колонок в master.csv (не state.json)
  - Предлагаем загрузку если найдены новые периоды
  - Парсинг запускается только после подтверждения
"""

import asyncio
import logging
from datetime import datetime

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from config.settings import MONTH_RU, PERIOD_ID_MONTHLY, PERIOD_ID_YEARLY
from core.periods import (
    period_sort_key, fmt_period, period_key, period_label,
    yearly_key, is_yearly_key, is_monthly_key, parse_period_key,
)
from core.storage import get_cookie, get_master_period_cols
from core.parser import check_session_both
from core.tasks import run_parse_and_update

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# ВСПОМОГАТЕЛЬНЫЕ
# ─────────────────────────────────────────────

def _no_cookie_reply(q):
    return q.edit_message_text(
        "❌ Cookie не установлены.\nНажми 🔑 Обновить Cookie",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("🔑 Обновить Cookie", callback_data="set_cookie")
        ]])
    )


def _bad_session_reply(q):
    return q.edit_message_text(
        "❌ Сессия недействительна.\n"
        "Обнови Cookie: F12 → Network → заголовок Cookie",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("🔑 Обновить Cookie", callback_data="set_cookie")
        ]])
    )


async def _detect_new_periods(
    cookie: str,
) -> tuple[bool, list[str], list[str]]:
    """
    Fetches available periods from the website and compares against
    what is already in master.csv. Returns (ok, new_monthly, new_yearly).
    """
    ok, monthly, yearly = await asyncio.get_event_loop().run_in_executor(
        None, check_session_both, cookie
    )
    if not ok:
        return False, [], []

    local_cols = get_master_period_cols()

    new_monthly = sorted(
        [p for p in monthly if p not in local_cols and is_monthly_key(p)],
        key=period_sort_key,
    )
    new_yearly = sorted(
        [p for p in yearly if p not in local_cols],
        key=period_sort_key,
    )
    return True, new_monthly, new_yearly


def _fmt_period_list(periods: list[str], limit: int = 6) -> list[str]:
    lines = [f"  • {fmt_period(p)}" for p in periods[:limit]]
    if len(periods) > limit:
        lines.append(f"  ...и ещё {len(periods) - limit}")
    return lines


def _start_parse_task(q, ctx, pm: dict, new_periods: list[str], data_type: str) -> None:
    """Launches the parse task after confirmation."""
    type_label = "годовых" if data_type == "yearly" else "месячных"
    lines = [f"🆕 Загружаю {type_label} периодов: *{len(new_periods)}*"]
    lines += _fmt_period_list(new_periods)
    lines.append("\n⏳ Начинаю загрузку...")
    asyncio.ensure_future(_edit_and_run(q, ctx, pm, lines, data_type))


async def _edit_and_run(q, ctx, pm: dict, lines: list[str], data_type: str) -> None:
    await q.edit_message_text("\n".join(lines), parse_mode="Markdown")
    cookie = get_cookie()
    asyncio.create_task(
        run_parse_and_update(q.message.chat_id, ctx.bot, cookie, pm, data_type=data_type)
    )


# ─────────────────────────────────────────────
# ПРОВЕРКА НОВЫХ ДАННЫХ
# ─────────────────────────────────────────────

async def cb_check(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Проверяет сайт на новые периоды, сравнивая с колонками в master.csv.
    Показывает статус и предлагает загрузку если есть новые данные.
    """
    q = update.callback_query
    await q.answer()

    cookie = get_cookie()
    if not cookie:
        await _no_cookie_reply(q)
        return

    await q.edit_message_text("🔍 Проверяю доступные данные на сайте...")

    ok, new_monthly, new_yearly = await _detect_new_periods(cookie)
    if not ok:
        await _bad_session_reply(q)
        return

    # Store for confirm callbacks
    ctx.user_data["new_monthly_periods"] = new_monthly
    ctx.user_data["new_yearly_periods"] = new_yearly

    local_cols = get_master_period_cols()
    local_monthly = sorted([c for c in local_cols if is_monthly_key(c)], key=period_sort_key)
    local_yearly  = sorted([c for c in local_cols if is_yearly_key(c)],  key=period_sort_key)

    lines = ["✅ Сессия активна.\n"]
    if local_monthly:
        lines.append(f"📅 В базе месячных: {fmt_period(local_monthly[0])} — {fmt_period(local_monthly[-1])}")
    if local_yearly:
        lines.append(f"📆 В базе годовых: {fmt_period(local_yearly[0])} — {fmt_period(local_yearly[-1])}")

    if new_monthly:
        lines.append(f"\n🆕 *Новых месячных на сайте: {len(new_monthly)}*")
        lines += _fmt_period_list(new_monthly)
    if new_yearly:
        lines.append(f"\n🆕 *Новых годовых на сайте: {len(new_yearly)}*")
        lines += _fmt_period_list(new_yearly)
    if not new_monthly and not new_yearly:
        lines.append("\n✅ Все данные уже актуальны.")

    kb = []
    if new_monthly:
        kb.append([InlineKeyboardButton(
            f"⬇️ Загрузить месячные ({len(new_monthly)})",
            callback_data="confirm_auto_monthly",
        )])
    if new_yearly:
        kb.append([InlineKeyboardButton(
            f"📆 Загрузить годовые ({len(new_yearly)})",
            callback_data="confirm_auto_yearly",
        )])
    kb.append([InlineKeyboardButton("🏠 Главное меню", callback_data="iok_menu")])

    await q.edit_message_text(
        "\n".join(lines),
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb),
    )


# ─────────────────────────────────────────────
# ОБНОВЛЕНИЕ ДАННЫХ (ПРОВЕРКА + ПРЕДЛОЖЕНИЕ)
# ─────────────────────────────────────────────

async def cb_update_auto(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Проверяет сайт на новые периоды (месячные + годовые) и предлагает
    загрузить найденные. Сравнение ведётся против колонок в master.csv.
    """
    q = update.callback_query
    await q.answer()

    cookie = get_cookie()
    if not cookie:
        await _no_cookie_reply(q)
        return

    await q.edit_message_text("🔍 Проверяю новые данные на сайте...")

    ok, new_monthly, new_yearly = await _detect_new_periods(cookie)
    if not ok:
        await _bad_session_reply(q)
        return

    # Store for confirm callbacks
    ctx.user_data["new_monthly_periods"] = new_monthly
    ctx.user_data["new_yearly_periods"] = new_yearly

    if not new_monthly and not new_yearly:
        await q.edit_message_text(
            "✅ Данные ИОК уже актуальны — новых периодов на сайте не найдено.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🏠 Меню", callback_data="iok_menu")
            ]])
        )
        return

    lines = ["📊 *Найдены новые данные ИОК на сайте:*\n"]
    if new_monthly:
        lines.append(f"📅 Новых месячных: *{len(new_monthly)}*")
        lines += _fmt_period_list(new_monthly)
    if new_yearly:
        lines.append(f"\n📆 Новых годовых: *{len(new_yearly)}*")
        lines += _fmt_period_list(new_yearly)
    lines.append("\nЗагрузка займёт 30–90 минут.")

    kb = []
    if new_monthly:
        kb.append([InlineKeyboardButton(
            f"⬇️ Загрузить месячные ({len(new_monthly)})",
            callback_data="confirm_auto_monthly",
        )])
    if new_yearly:
        kb.append([InlineKeyboardButton(
            f"📆 Загрузить годовые ({len(new_yearly)})",
            callback_data="confirm_auto_yearly",
        )])
    kb.append([InlineKeyboardButton("❌ Отмена", callback_data="iok_menu")])

    await q.edit_message_text(
        "\n".join(lines),
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb),
    )


# ─────────────────────────────────────────────
# ПОДТВЕРЖДЕНИЕ И ЗАПУСК
# ─────────────────────────────────────────────

async def cb_confirm_auto_monthly(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Запускает парсинг новых месячных периодов."""
    q = update.callback_query
    await q.answer()

    new_periods = ctx.user_data.get("new_monthly_periods", [])
    if not new_periods:
        await q.edit_message_text(
            "❌ Список периодов устарел. Нажми «Загрузить новые данные» снова.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🏠 Меню", callback_data="iok_menu")
            ]])
        )
        return

    pm = {p: period_label(int(p[1:3]), int(p[3:7])) for p in new_periods}
    _start_parse_task(q, ctx, pm, new_periods, "monthly")


async def cb_confirm_auto_yearly(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Запускает парсинг новых годовых периодов."""
    q = update.callback_query
    await q.answer()

    new_periods = ctx.user_data.get("new_yearly_periods", [])
    if not new_periods:
        await q.edit_message_text(
            "❌ Список периодов устарел. Нажми «Загрузить новые данные» снова.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🏠 Меню", callback_data="iok_menu")
            ]])
        )
        return

    pm = {p: fmt_period(p) for p in new_periods}
    _start_parse_task(q, ctx, pm, new_periods, "yearly")


# ─────────────────────────────────────────────
# РУЧНОЙ ВЫБОР: МЕСЯЧНЫЕ ДАННЫЕ
# ─────────────────────────────────────────────

async def cb_manual_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Показывает выбор года для месячных данных."""
    q = update.callback_query
    await q.answer()
    ctx.user_data["data_type"] = "monthly"
    cur = datetime.now().year
    kb = [
        [InlineKeyboardButton(str(y), callback_data=f"sel_year_{y}")]
        for y in range(cur, 2018, -1)
    ]
    kb.append([InlineKeyboardButton("❌ Отмена", callback_data="iok_menu")])
    await q.edit_message_text(
        "📅 *Загрузка помесячных данных*\nВыбери год:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb)
    )


async def cb_sel_year(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Показывает выбор месяца для выбранного года."""
    q = update.callback_query
    await q.answer()
    year = int(q.data.split("_")[-1])
    ctx.user_data["sel_year"] = year

    kb: list = []
    row: list = []
    for m in range(1, 13):
        row.append(InlineKeyboardButton(MONTH_RU[m], callback_data=f"sel_month_{m}"))
        if len(row) == 4:
            kb.append(row)
            row = []
    if row:
        kb.append(row)

    kb.append([InlineKeyboardButton("📅 Весь год (янв-дек)", callback_data="sel_month_0")])
    kb.append([InlineKeyboardButton("◀️ Назад", callback_data="manual_start")])

    await q.edit_message_text(
        f"📅 Год: *{year}*\n"
        f"Выбери последний месяц периода (загрузится Янв—X):",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb)
    )


async def cb_sel_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Подтверждение выбранного периода."""
    q = update.callback_query
    await q.answer()
    month = int(q.data.split("_")[-1])
    year = ctx.user_data.get("sel_year", datetime.now().year)

    months_range = list(range(1, 13)) if month == 0 else list(range(1, month + 1))
    label = f"весь {year}" if month == 0 else period_label(month, year)

    pm = {period_key(m, year): period_label(m, year) for m in months_range}
    ctx.user_data["period_map"] = pm
    ctx.user_data["sel_label"] = label
    ctx.user_data["data_type"] = "monthly"

    local_cols = get_master_period_cols()
    already = [p for p in pm if p in local_cols]
    new_only = [p for p in pm if p not in local_cols]

    lines = [f"📅 Выбран период: *{label}*\n", f"Всего периодов: {len(pm)}"]
    if already:
        lines.append(f"✅ Уже в базе: {len(already)} (пустые ячейки будут заполнены)")
    if new_only:
        lines.append(f"🆕 Новых для загрузки: {len(new_only)}")
    else:
        lines.append("ℹ️ Все эти периоды уже в базе (пустые ячейки будут заполнены)")
    lines.append("\n⏳ Загрузка займёт 30–90 минут")

    kb = [
        [InlineKeyboardButton("✅ Загрузить", callback_data="confirm_manual")],
        [InlineKeyboardButton("◀️ Назад", callback_data=f"sel_year_{year}")],
    ]
    await q.edit_message_text(
        "\n".join(lines), parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb)
    )


# ─────────────────────────────────────────────
# РУЧНОЙ ВЫБОР: ГОДОВЫЕ ДАННЫЕ
# ─────────────────────────────────────────────

async def cb_yearly_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Показывает выбор годов для загрузки годовых данных."""
    q = update.callback_query
    await q.answer()
    ctx.user_data["data_type"] = "yearly"
    cur = datetime.now().year

    local_cols = get_master_period_cols()

    kb: list = []
    row: list = []
    for y in range(cur, 2009, -1):
        yk = yearly_key(y)
        mark = "✅" if yk in local_cols else ""
        row.append(InlineKeyboardButton(f"{mark}{y}", callback_data=f"sel_yearly_{y}"))
        if len(row) == 3:
            kb.append(row)
            row = []
    if row:
        kb.append(row)

    kb.append([InlineKeyboardButton("📆 Все доступные годы", callback_data="sel_yearly_all")])
    kb.append([InlineKeyboardButton("❌ Отмена", callback_data="iok_menu")])

    await q.edit_message_text(
        "📆 *Загрузка годовых данных*\n"
        "Выбери год (✅ = уже загружен):",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb)
    )


async def cb_sel_yearly(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Подтверждение выбранного года для годовых данных."""
    q = update.callback_query
    await q.answer()

    data = q.data
    if data == "sel_yearly_all":
        cur = datetime.now().year
        years = list(range(2010, cur + 1))
        label = f"все годы (2010-{cur})"
    else:
        y = int(data.split("_")[-1])
        years = [y]
        label = str(y)

    pm = {yearly_key(y): str(y) for y in years}
    ctx.user_data["period_map"] = pm
    ctx.user_data["sel_label"] = label
    ctx.user_data["data_type"] = "yearly"

    local_cols = get_master_period_cols()
    already = [p for p in pm if p in local_cols]
    new_only = [p for p in pm if p not in local_cols]

    lines = [f"📆 Годовые данные: *{label}*\n"]
    if already:
        lines.append(f"✅ Уже в базе: {len(already)}")
    if new_only:
        lines.append(f"🆕 Новых: {len(new_only)}")
    lines.append("\n⏳ Загрузка займёт 30–90 минут")

    kb = [
        [InlineKeyboardButton("✅ Загрузить", callback_data="confirm_manual")],
        [InlineKeyboardButton("◀️ Назад", callback_data="yearly_start")],
    ]
    await q.edit_message_text(
        "\n".join(lines), parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb)
    )


# ─────────────────────────────────────────────
# ПОДТВЕРЖДЕНИЕ (РУЧНОЙ ВЫБОР)
# ─────────────────────────────────────────────

async def cb_confirm_manual(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Запускает загрузку выбранного вручную периода."""
    q = update.callback_query
    await q.answer()

    pm = ctx.user_data.get("period_map", {})
    label = ctx.user_data.get("sel_label", "выбранный период")
    data_type = ctx.user_data.get("data_type", "monthly")
    cookie = get_cookie()

    if not pm:
        await q.edit_message_text("❌ Периоды не выбраны. Попробуй снова /start")
        return

    if not cookie:
        await _no_cookie_reply(q)
        return

    type_label = "годовых" if data_type == "yearly" else "месячных"
    await q.edit_message_text(
        f"⏳ Начинаю загрузку {type_label} данных за *{label}*...\n"
        f"Пришлю Excel когда закончу.",
        parse_mode="Markdown",
    )

    asyncio.create_task(
        run_parse_and_update(q.message.chat_id, ctx.bot, cookie, pm, data_type=data_type)
    )


# ─────────────────────────────────────────────
# RESUME — ПРОДОЛЖИТЬ ПРЕРВАННЫЙ ПАРСИНГ
# ─────────────────────────────────────────────

async def cb_resume(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Продолжает прерванный парсинг если есть файл прогресса."""
    q = update.callback_query
    await q.answer()

    from config.settings import DATA_DIR
    import pandas as pd
    from core.periods import is_period_key

    progress_path = DATA_DIR / "parse_progress.csv"

    if not progress_path.exists():
        await q.edit_message_text(
            "ℹ️ Нет прерванного парсинга для продолжения.\n"
            "Используй 🔍 Проверить новые данные или 📅 Загрузить помесячно.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🏠 Меню", callback_data="iok_menu")
            ]])
        )
        return

    cookie = get_cookie()
    if not cookie:
        await _no_cookie_reply(q)
        return

    try:
        df = pd.read_csv(progress_path, encoding="utf-8-sig", low_memory=False)
        done_regions = len(set(zip(df["oblast"].astype(str), df["region"].astype(str))))
        progress_rows = len(df)
        period_cols = [c for c in df.columns if is_period_key(c)]
    except Exception:
        await q.edit_message_text(
            "❌ Не удалось прочитать файл прогресса.\nСбрось прогресс и начни заново.",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🗑 Сбросить прогресс", callback_data="reset_progress")],
                [InlineKeyboardButton("🏠 Меню", callback_data="iok_menu")],
            ])
        )
        return

    from config.regions import ALL_REGIONS
    total = sum(len(v) for v in ALL_REGIONS.values())

    kb = [
        [InlineKeyboardButton("▶️ Продолжить парсинг", callback_data="do_resume")],
        [InlineKeyboardButton("🗑 Сбросить прогресс", callback_data="reset_progress")],
        [InlineKeyboardButton("🏠 Меню", callback_data="iok_menu")],
    ]

    await q.edit_message_text(
        f"🔄 *Найден прерванный парсинг*\n\n"
        f"Пройдено районов: {done_regions}/{total}\n"
        f"Строк в прогрессе: {progress_rows:,}\n"
        f"Периодов: {len(period_cols)}\n\n"
        f"Продолжить с места остановки?",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb),
    )


async def cb_do_resume(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Запускает продолжение парсинга."""
    q = update.callback_query
    await q.answer()

    from config.settings import DATA_DIR
    import pandas as pd
    from core.periods import is_period_key, parse_period_key

    cookie = get_cookie()
    progress_path = DATA_DIR / "parse_progress.csv"

    period_cols = []
    if progress_path.exists():
        try:
            df = pd.read_csv(progress_path, encoding="utf-8-sig", nrows=1)
            period_cols = [c for c in df.columns if is_period_key(c)]
        except Exception:
            pass

    if not period_cols:
        await q.edit_message_text(
            "❌ Не удалось определить периоды из прогресса.\nНачни парсинг заново.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🏠 Меню", callback_data="iok_menu")
            ]])
        )
        return

    pm = {}
    for c in period_cols:
        parsed = parse_period_key(c)
        if parsed:
            pm[c] = period_label(*parsed)

    has_yearly = any(is_yearly_key(c) for c in period_cols)
    data_type = "yearly" if has_yearly else "monthly"

    await q.edit_message_text(
        f"⏳ Продолжаю парсинг...\n"
        f"Периодов: {len(pm)}\n"
        f"Пришлю Excel когда закончу.",
        parse_mode="Markdown",
    )

    asyncio.create_task(
        run_parse_and_update(q.message.chat_id, ctx.bot, cookie, pm, data_type=data_type)
    )


async def cb_reset_progress(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Сбрасывает файл прогресса."""
    q = update.callback_query
    await q.answer()

    from config.settings import DATA_DIR

    for fname in ("parse_progress.csv", "parse_done.txt"):
        try:
            (DATA_DIR / fname).unlink(missing_ok=True)
        except Exception:
            pass

    await q.edit_message_text(
        "🗑 Прогресс сброшен.",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("🏠 Меню", callback_data="iok_menu")
        ]])
    )
