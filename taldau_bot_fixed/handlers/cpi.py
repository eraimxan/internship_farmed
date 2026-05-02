"""
handlers/cpi.py — Telegram-хендлеры для раздела ИПЦ (инфляция).
"""

import asyncio
import logging
import re
from datetime import datetime

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from core.storage import get_cookie
from ai_analyzer import _period_sort_key, analyze_cpi
from core.cpi_tasks import (
    _existing, PERIOD_LABELS, MONTH_NAMES, QUARTER_MONTHS,
    get_available_years, get_available_months, build_filtered_excel,
    run_cpi_pipeline, check_cpi_new_periods,
)
from config.settings import TELEGRAM_FILE_LIMIT

logger = logging.getLogger(__name__)

CPI_MENU_KB = [
    [InlineKeyboardButton("🔍 Проверить и получить новые данные", callback_data="cpi_check_upd")],
    [InlineKeyboardButton("📥 Получить файл Excel",               callback_data="cpi_download")],
    [InlineKeyboardButton("🔁 Загрузить всё заново",              callback_data="cpi_fetch_fresh")],
    [InlineKeyboardButton("📊 Состояние данных",                  callback_data="cpi_status")],
    [InlineKeyboardButton("← Назад",                             callback_data="main_menu")],
]

_BACK_CPI   = [[InlineKeyboardButton("◀️ Назад в меню ИПЦ", callback_data="cpi_menu")]]
_BACK_DLOPT = [[InlineKeyboardButton("◀️ Назад",            callback_data="cpi_dl_back")]]


# ── Выбор типа периода ────────────────────────────────────────────────────────

def _period_kb(prefix: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📅 Месяц",       callback_data=f"{prefix}4")],
        [InlineKeyboardButton("📊 Квартал",     callback_data=f"{prefix}5")],
        [InlineKeyboardButton("📈 Накопленный", callback_data=f"{prefix}9")],
        [InlineKeyboardButton("◀️ Назад",       callback_data="cpi_menu")],
    ])


def _fresh_filter_kb(period_id: str) -> InlineKeyboardMarkup:
    """Filter options shown after choosing a period type in 'Download again'."""
    is_quarterly = period_id == "5"
    rows = [
        [InlineKeyboardButton("🔁 Все данные заново",     callback_data="cpi_df_all")],
        [InlineKeyboardButton("📅 Выбрать год",           callback_data="cpi_df_yr")],
    ]
    if is_quarterly:
        rows.append([InlineKeyboardButton("📊 Выбрать квартал", callback_data="cpi_df_qr")])
    else:
        rows.append([InlineKeyboardButton("🗓 Выбрать месяц",   callback_data="cpi_df_mo")])
    rows.append([InlineKeyboardButton("◀️ Назад", callback_data="cpi_fetch_fresh")])
    return InlineKeyboardMarkup(rows)


# ── Главное меню ИПЦ ──────────────────────────────────────────────────────────

def _cpi_brief_status() -> str:
    lines = []
    for pid, label in PERIOD_LABELS.items():
        csv_f = _existing(pid, "csv")
        xl_f  = _existing(pid, "excel")
        if csv_f or xl_f:
            dt = ""
            if xl_f:
                dt = datetime.fromtimestamp(xl_f.stat().st_mtime).strftime("%d.%m.%Y")
            lines.append(f"  • {label}: {'обновлён ' + dt if dt else 'есть CSV'}")
    if not lines:
        return "📭 Данных пока нет — нажми *Проверить и получить новые данные*"
    return "📦 Загруженные данные:\n" + "\n".join(lines)


async def cb_cpi_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    await q.edit_message_text(
        f"*📈 ИПЦ — Индексы потребительских цен*\n\n"
        f"Данные об инфляции по 21 региону Казахстана.\n\n"
        f"{_cpi_brief_status()}\n\n"
        f"Что хотите сделать?",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(CPI_MENU_KB),
    )


# ── Проверка обновлений ───────────────────────────────────────────────────────

async def cb_cpi_check_upd(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    await q.edit_message_text(
        "*🔍 Проверка обновлений — ИПЦ*\n\n"
        "Выберите тип данных для проверки:\n\n"
        "📅 *Месяц* — данные за каждый отдельный месяц\n"
        "📊 *Квартал* — данные за каждые 3 месяца\n"
        "📈 *Накопленный* — нарастающим итогом с начала года",
        parse_mode="Markdown",
        reply_markup=_period_kb("cpi_cu"),
    )


async def cb_cpi_do_check(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Checks for new periods vs local CSV.
    When found: shows the list and ONE button to download all new.
    Handles cpi_cu*.
    """
    q = update.callback_query
    period_id = q.data[-1]

    cookie = get_cookie()
    if not cookie:
        await q.answer()
        await q.edit_message_text(
            "❌ *Не установлен Cookie*\n\n"
            "Без Cookie бот не может зайти на сайт taldau.stat.gov.kz.\n\n"
            "Как получить Cookie:\n"
            "1. Открой сайт taldau.stat.gov.kz\n"
            "2. Нажми F12 → Network → обнови страницу\n"
            "3. Найди запрос `GetIndexTreeData` → скопируй `Cookie:`\n"
            "4. Вставь ниже через кнопку",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🔑 Вставить Cookie", callback_data="set_cookie"),
                InlineKeyboardButton("◀️ Назад",           callback_data="cpi_menu"),
            ]]),
        )
        return

    label = PERIOD_LABELS.get(period_id, period_id)
    await q.answer()
    await q.edit_message_text(
        f"🔍 Проверяю новые данные — *{label}*...\n"
        f"Сравниваю с сайтом. Обычно 10–30 секунд.",
        parse_mode="Markdown",
    )

    ok, new_keys, period_map = await check_cpi_new_periods(cookie, period_id)

    if not ok:
        await q.edit_message_text(
            "❌ Не удалось подключиться к сайту. Проверь Cookie.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🔑 Вставить Cookie", callback_data="set_cookie"),
                InlineKeyboardButton("◀️ Назад",           callback_data="cpi_menu"),
            ]]),
        )
        return

    if not new_keys:
        await q.edit_message_text(
            f"✅ *{label}* — данные уже актуальны.\nНовых периодов на сайте не найдено.",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("◀️ Меню ИПЦ", callback_data="cpi_menu"),
            ]]),
        )
        return

    sorted_keys = sorted(new_keys, key=_period_sort_key)

    ctx.user_data["cpi_new_keys"]  = sorted_keys
    ctx.user_data["cpi_new_map"]   = period_map
    ctx.user_data["cpi_check_pid"] = period_id

    names_text = "\n".join(f"  • {period_map.get(k, k)}" for k in sorted_keys)
    await q.edit_message_text(
        f"📊 Найдено *{len(new_keys)}* новых период(ов) — *{label}*:\n\n"
        f"{names_text}\n\n"
        f"Нажмите кнопку чтобы загрузить все новые данные:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton(
                f"⬇️ Загрузить новые данные ({len(sorted_keys)})",
                callback_data="cpi_sel_all",
            )],
            [InlineKeyboardButton("❌ Отмена", callback_data="cpi_menu")],
        ]),
    )


async def cb_cpi_sel_period(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Handles cpi_sel_all — downloads all newly found periods.
    (Individual period selection removed; this only handles 'all'.)
    """
    q = update.callback_query

    period_id  = ctx.user_data.get("cpi_check_pid", "5")
    new_keys   = ctx.user_data.get("cpi_new_keys", [])
    period_map = ctx.user_data.get("cpi_new_map", {})
    label      = PERIOD_LABELS.get(period_id, period_id)

    cookie = get_cookie()
    if not cookie:
        await q.answer("❌ Cookie не установлены")
        return

    if not new_keys:
        await q.answer("❌ Список устарел — нажми «Проверить» снова")
        await q.edit_message_text(
            "❌ Данные о новых периодах устарели. Запусти проверку заново.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("◀️ Меню ИПЦ", callback_data="cpi_menu"),
            ]]),
        )
        return

    analysis_keys = set(new_keys)

    await q.answer()
    await q.edit_message_text(
        f"⏳ Начинаю загрузку — *{label}*\n"
        f"Новых периодов: *{len(new_keys)}*\n\n"
        f"Скачиваю данные по 21 региону. Займёт 5–15 минут.\n"
        f"Пришлю Excel и аналитику когда закончу.",
        parse_mode="Markdown",
    )

    asyncio.create_task(
        run_cpi_pipeline(
            q.message.chat_id, ctx.bot, cookie,
            period_id=period_id,
            no_resume=True,
            analysis_period_keys=analysis_keys,
        )
    )


# ── Загрузка заново ───────────────────────────────────────────────────────────

async def cb_cpi_fetch_fresh(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    ctx.user_data.pop("cpi_dl_fresh", None)
    await q.edit_message_text(
        "*🔁 Загрузить всё заново — ИПЦ*\n\n"
        "Скачивает данные с нуля с сайта.\n\n"
        "Выберите тип данных:\n\n"
        "📅 *Месяц* — данные за каждый отдельный месяц\n"
        "📊 *Квартал* — данные за каждые 3 месяца\n"
        "📈 *Накопленный* — нарастающим итогом с начала года",
        parse_mode="Markdown",
        reply_markup=_period_kb("cpi_nr"),
    )


async def cb_cpi_do_fetch(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Period type selected for 'Download again'. Shows filter options.
    Handles cpi_nr*.
    """
    q = update.callback_query
    period_id = q.data[-1]
    label     = PERIOD_LABELS.get(period_id, period_id)

    ctx.user_data["cpi_dl_pid"]   = period_id
    ctx.user_data["cpi_dl_fresh"] = True

    cookie = get_cookie()
    if not cookie:
        await q.answer()
        await q.edit_message_text(
            "❌ *Не установлен Cookie*\n\nОбнови Cookie через настройки.",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("◀️ Назад", callback_data="cpi_menu"),
            ]]),
        )
        return

    has_data = bool(get_available_years(period_id))

    await q.answer()
    await q.edit_message_text(
        f"*🔁 Загрузить заново — {label}*\n\n"
        f"Выберите за какой период нужен Excel.\n"
        f"Данные всегда скачиваются полностью с сайта.\n\n"
        + ("" if has_data else "_(Для выбора периода нужны уже загруженные данные. Если данных нет — выбери «Все данные заново».)_\n\n"),
        parse_mode="Markdown",
        reply_markup=_fresh_filter_kb(period_id),
    )


# ── Скачать Excel ─────────────────────────────────────────────────────────────

async def cb_cpi_download(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()
    ctx.user_data["cpi_dl_fresh"] = False
    await q.edit_message_text(
        "*📥 Получить файл Excel — ИПЦ*\n\n"
        "Выберите тип данных:\n\n"
        "📅 *Месяц* — данные за каждый отдельный месяц\n"
        "📊 *Квартал* — данные за каждые 3 месяца\n"
        "📈 *Накопленный* — нарастающим итогом с начала года",
        parse_mode="Markdown",
        reply_markup=_period_kb("cpi_dp"),
    )


async def cb_cpi_do_download(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Period type selected for 'Get Excel'. Shows filter options. Handles cpi_dp*."""
    q = update.callback_query
    period_id = q.data[-1]
    ctx.user_data["cpi_dl_pid"]   = period_id
    ctx.user_data["cpi_dl_fresh"] = False
    label = PERIOD_LABELS.get(period_id, period_id)
    is_quarterly = period_id == "5"
    await q.answer()
    await q.edit_message_text(
        f"*📥 Получить Excel — {label}*\n\n"
        f"За какой период нужны данные?",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("📋 Все данные",          callback_data="cpi_df_all")],
            [InlineKeyboardButton("📅 За конкретный год",   callback_data="cpi_df_yr")],
            [InlineKeyboardButton("📊 За квартал",          callback_data="cpi_df_qr")] if is_quarterly else
            [InlineKeyboardButton("🗓 За конкретный месяц", callback_data="cpi_df_mo")],
            [InlineKeyboardButton("◀️ Назад",               callback_data="cpi_download")],
        ]),
    )


async def cb_cpi_dl_back(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Return to filter selection (preserves period type and fresh flag)."""
    q     = update.callback_query
    pid   = ctx.user_data.get("cpi_dl_pid", "5")
    label = PERIOD_LABELS.get(pid, pid)
    fresh = ctx.user_data.get("cpi_dl_fresh", False)
    is_quarterly = pid == "5"
    await q.answer()
    if fresh:
        await q.edit_message_text(
            f"*🔁 Загрузить заново — {label}*\n\nВыберите период:",
            parse_mode="Markdown",
            reply_markup=_fresh_filter_kb(pid),
        )
    else:
        await q.edit_message_text(
            f"*📥 Получить Excel — {label}*\n\nЗа какой период нужны данные?",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("📋 Все данные",          callback_data="cpi_df_all")],
                [InlineKeyboardButton("📅 За конкретный год",   callback_data="cpi_df_yr")],
                [InlineKeyboardButton("📊 За квартал",          callback_data="cpi_df_qr")] if is_quarterly else
                [InlineKeyboardButton("🗓 За конкретный месяц", callback_data="cpi_df_mo")],
                [InlineKeyboardButton("◀️ Назад",               callback_data="cpi_download")],
            ]),
        )


# ── Выбор фильтра времени ─────────────────────────────────────────────────────

async def cb_cpi_dl_filter(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Handles cpi_df_all / cpi_df_yr / cpi_df_mo / cpi_df_qr."""
    q    = update.callback_query
    data = q.data
    pid  = ctx.user_data.get("cpi_dl_pid", "5")
    fresh = ctx.user_data.get("cpi_dl_fresh", False)
    await q.answer()

    if data == "cpi_df_all":
        if fresh:
            # Full re-download, no filter
            await _launch_fresh(q, ctx, pid)
        else:
            await _send_filtered(q, ctx, pid, "all")
        return

    mode_map = {"cpi_df_yr": "year", "cpi_df_mo": "month", "cpi_df_qr": "quarter"}
    mode = mode_map[data]
    ctx.user_data["cpi_dl_mode"] = mode

    years = get_available_years(pid)
    if not years:
        if fresh:
            await q.edit_message_text(
                "❌ Нет данных для выбора периода.\n\n"
                "Нажми *Все данные заново* чтобы сначала загрузить.",
                parse_mode="Markdown",
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton("◀️ Назад", callback_data="cpi_dl_back"),
                ]]),
            )
        else:
            await q.edit_message_text(
                "❌ Данных нет.\n\nСначала загрузи данные через *Проверить и получить новые данные*.",
                parse_mode="Markdown",
                reply_markup=InlineKeyboardMarkup(_BACK_CPI),
            )
        return

    prefix = {"year": "cpi_yf_", "month": "cpi_ym_", "quarter": "cpi_yq_"}[mode]
    kb = [[InlineKeyboardButton(y, callback_data=f"{prefix}{y}")] for y in years]
    kb.append([InlineKeyboardButton("◀️ Назад", callback_data="cpi_dl_back")])

    await q.edit_message_text(
        "📅 *Выберите год:*",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(kb),
    )


# ── Год выбран ────────────────────────────────────────────────────────────────

async def cb_cpi_dl_year(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q    = update.callback_query
    data = q.data
    year = data.split("_")[-1]
    pid  = ctx.user_data.get("cpi_dl_pid", "5")
    fresh = ctx.user_data.get("cpi_dl_fresh", False)
    ctx.user_data["cpi_dl_year"] = year
    await q.answer()

    if data.startswith("cpi_yf_"):
        # Year selected — build Excel for whole year (+ AI)
        if fresh:
            await _launch_fresh(q, ctx, pid, filter_type="year", filter_year=year)
        else:
            await _send_filtered(q, ctx, pid, "year", year=year)

    elif data.startswith("cpi_ym_"):
        months = get_available_months(pid, year)
        if not months:
            await q.edit_message_text(
                f"❌ Нет данных за {year}.",
                reply_markup=InlineKeyboardMarkup(_BACK_DLOPT),
            )
            return
        kb = [
            [InlineKeyboardButton(MONTH_NAMES.get(int(m), m), callback_data=f"cpi_mo_{m}")]
            for m in months
        ]
        kb.append([InlineKeyboardButton("◀️ Назад", callback_data="cpi_dl_back")])
        await q.edit_message_text(
            f"🗓 *Выберите месяц ({year}):*",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(kb),
        )

    elif data.startswith("cpi_yq_"):
        kb = [
            [InlineKeyboardButton(
                f"Q{q_num} — {MONTH_NAMES[int(QUARTER_MONTHS[q_num][0])]}–{MONTH_NAMES[int(QUARTER_MONTHS[q_num][-1])]}",
                callback_data=f"cpi_qu_{q_num}",
            )]
            for q_num in [1, 2, 3, 4]
        ]
        kb.append([InlineKeyboardButton("◀️ Назад", callback_data="cpi_dl_back")])
        await q.edit_message_text(
            f"📊 *Выберите квартал ({year}):*",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(kb),
        )


# ── Месяц / Квартал выбран ────────────────────────────────────────────────────

async def cb_cpi_dl_month(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q     = update.callback_query
    month = q.data.split("_")[-1]
    pid   = ctx.user_data.get("cpi_dl_pid", "5")
    year  = ctx.user_data.get("cpi_dl_year", "")
    fresh = ctx.user_data.get("cpi_dl_fresh", False)
    await q.answer()
    if fresh:
        await _launch_fresh(q, ctx, pid, filter_type="month", filter_year=year, filter_month=month)
    else:
        await _send_filtered(q, ctx, pid, "month", year=year, month=month)


async def cb_cpi_dl_quarter(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q       = update.callback_query
    quarter = int(q.data.split("_")[-1])
    pid     = ctx.user_data.get("cpi_dl_pid", "5")
    year    = ctx.user_data.get("cpi_dl_year", "")
    fresh   = ctx.user_data.get("cpi_dl_fresh", False)
    await q.answer()
    if fresh:
        await _launch_fresh(q, ctx, pid, filter_type="quarter", filter_year=year, filter_quarter=quarter)
    else:
        await _send_filtered(q, ctx, pid, "quarter", year=year, quarter=quarter)


# ── Helpers ───────────────────────────────────────────────────────────────────

async def _launch_fresh(
    q, ctx,
    pid: str,
    filter_type: str = "all",
    filter_year: str = None,
    filter_month: str = None,
    filter_quarter: int = None,
) -> None:
    """Start a full re-fetch pipeline with optional filter for Excel + AI."""
    label  = PERIOD_LABELS.get(pid, pid)
    cookie = get_cookie()
    if not cookie:
        await q.edit_message_text(
            "❌ Cookie не установлены.",
            reply_markup=InlineKeyboardMarkup(_BACK_CPI),
        )
        return

    if filter_type == "month" and filter_month and filter_year:
        desc = MONTH_NAMES.get(int(filter_month), filter_month) + f" {filter_year}"
    elif filter_type == "year" and filter_year:
        desc = filter_year
    elif filter_type == "quarter" and filter_quarter and filter_year:
        desc = f"Q{filter_quarter} {filter_year}"
    else:
        desc = "все данные"

    ctx.user_data.pop("cpi_dl_fresh", None)

    await q.edit_message_text(
        f"⏳ Начинаю загрузку — *{label}*\n"
        f"Период Excel: *{desc}*\n\n"
        f"Скачиваю данные по 21 региону. Займёт 5–15 минут.\n"
        f"Пришлю Excel и аналитику когда закончу.",
        parse_mode="Markdown",
    )

    asyncio.create_task(
        run_cpi_pipeline(
            q.message.chat_id, ctx.bot, cookie,
            period_id=pid,
            no_resume=True,
            filter_type=filter_type if filter_type != "all" else None,
            filter_year=filter_year,
            filter_month=filter_month,
            filter_quarter=filter_quarter,
        )
    )


async def _send_filtered(
    q, ctx,
    pid: str,
    filter_type: str,
    year: str = None,
    month: str = None,
    quarter: int = None,
) -> None:
    """Build filtered Excel from existing CSV, send it, then run AI analysis."""
    period_label = PERIOD_LABELS.get(pid, pid)
    await q.edit_message_text(
        f"⏳ Формирую файл Excel — {period_label}...\nОбычно занимает несколько секунд."
    )

    loop = asyncio.get_event_loop()
    try:
        excel_path, time_label = await loop.run_in_executor(
            None,
            lambda: build_filtered_excel(pid, filter_type, year=year, month=month, quarter=quarter),
        )
    except FileNotFoundError:
        await ctx.bot.send_message(
            q.message.chat_id,
            "❌ Данных нет.\n\nСначала загрузи данные через *Проверить и получить новые данные*.",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(_BACK_CPI),
        )
        return
    except ValueError as e:
        await ctx.bot.send_message(
            q.message.chat_id,
            f"❌ {e}",
            reply_markup=InlineKeyboardMarkup(_BACK_DLOPT),
        )
        return

    file_size = excel_path.stat().st_size
    sz_kb     = file_size // 1024

    if file_size > TELEGRAM_FILE_LIMIT:
        await ctx.bot.send_message(
            q.message.chat_id,
            f"⚠️ Файл слишком большой для отправки через Telegram ({sz_kb // 1024} МБ).\n"
            f"Файл сохранён на сервере:\n`{excel_path}`",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(_BACK_CPI),
        )
        return

    caption = f"📊 *ИПЦ — {period_label}*\n📅 Период: {time_label}\n📦 Размер: {sz_kb} КБ"
    with open(excel_path, "rb") as f:
        await ctx.bot.send_document(
            q.message.chat_id, f,
            filename=f"ИПЦ_{period_label}_{time_label.replace(' ', '_')}.xlsx",
            caption=caption,
            parse_mode="Markdown",
        )
    await ctx.bot.send_message(
        q.message.chat_id,
        "✅ Файл отправлен!",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("📥 Скачать ещё", callback_data="cpi_download"),
            InlineKeyboardButton("◀️ Меню ИПЦ",   callback_data="cpi_menu"),
        ]]),
    )

    # AI analysis for the selected period
    asyncio.create_task(
        _run_ai_for_filter(ctx.bot, q.message.chat_id, pid, filter_type, year, month, quarter)
    )


async def _run_ai_for_filter(
    bot, chat_id: int, pid: str,
    filter_type: str, year: str = None, month: str = None, quarter: int = None,
) -> None:
    """Runs AI analysis for the filtered period and sends result."""
    try:
        csv_path = _existing(pid, "csv")
        if not csv_path:
            return

        import pandas as pd
        df = pd.read_csv(csv_path, encoding="utf-8-sig", low_memory=False)
        all_cols = sorted(
            [c for c in df.columns if re.match(r"^y\d{6}$", c)],
            key=_period_sort_key,
        )

        if filter_type == "month" and month and year:
            target_cols = [c for c in all_cols if c == f"y{month}{year}"]
        elif filter_type == "year" and year:
            target_cols = [c for c in all_cols if c[3:] == year]
        elif filter_type == "quarter" and quarter and year:
            q_months = QUARTER_MONTHS[quarter]
            target_cols = [c for c in all_cols if c[1:3] in q_months and c[3:] == year]
        else:
            target_cols = all_cols[-3:]

        if not target_cols:
            return

        params_path = _existing(pid, "params")
        period_name_map: dict = {}
        if params_path:
            import json
            try:
                saved = json.loads(params_path.read_text(encoding="utf-8"))
                period_name_map = saved.get("period_name_map", {})
            except Exception:
                pass

        period_label = PERIOD_LABELS.get(pid, pid)
        labels = [period_name_map.get(c, c) for c in target_cols]

        loop = asyncio.get_event_loop()
        await bot.send_message(chat_id, "🤖 Генерирую аналитику...")
        analysis = await loop.run_in_executor(
            None, analyze_cpi, df, target_cols, labels, period_label
        )
        if analysis:
            try:
                await bot.send_message(chat_id, analysis, parse_mode="Markdown")
            except Exception:
                await bot.send_message(chat_id, analysis)
    except Exception as e:
        logger.warning(f"AI-анализ (фильтр) не удался: {e}")


# ── Статус данных ─────────────────────────────────────────────────────────────

async def cb_cpi_status(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    q = update.callback_query
    await q.answer()

    import pandas as pd
    lines = ["*📊 Состояние данных — ИПЦ*\n"]

    any_data = False
    for pid, label in PERIOD_LABELS.items():
        csv_f = _existing(pid, "csv")
        xl_f  = _existing(pid, "excel")
        ckpt  = _existing(pid, "checkpoint")

        if not csv_f and not xl_f:
            continue

        any_data = True
        lines.append(f"*{label}*")

        if csv_f:
            sz = csv_f.stat().st_size / 1024
            dt = datetime.fromtimestamp(csv_f.stat().st_mtime).strftime("%d.%m.%Y %H:%M")
            lines.append(f"  Данных: {sz:.0f} КБ | обновлено {dt}")
            try:
                df = pd.read_csv(csv_f, encoding="utf-8-sig", low_memory=False)
                regions = df["oblast"].nunique() if "oblast" in df.columns else "?"
                lines.append(f"  Строк: {len(df):,} | Регионов: {regions}")
            except Exception:
                pass

        if xl_f:
            sz = xl_f.stat().st_size / 1024
            dt = datetime.fromtimestamp(xl_f.stat().st_mtime).strftime("%d.%m.%Y %H:%M")
            lines.append(f"  Excel: {sz:.0f} КБ | {dt}")

        if ckpt:
            try:
                df_cp = pd.read_csv(ckpt, encoding="utf-8-sig", low_memory=False)
                done = df_cp["oblast"].nunique() if "oblast" in df_cp.columns else "?"
                lines.append(f"  ⚠️ Есть незавершённая загрузка: {done}/21 регионов")
            except Exception:
                pass

        lines.append("")

    if not any_data:
        lines.append("📭 Данных пока нет.\n\nНажми *Проверить и получить новые данные* чтобы начать.")

    await q.edit_message_text(
        "\n".join(lines).strip(),
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("◀️ Назад в меню ИПЦ", callback_data="cpi_menu"),
        ]]),
    )
