"""
handlers/download.py — Скачивание Excel и импорт CSV.
"""

import asyncio
import logging
import re

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from config.settings import OUTPUT_DIR, BASE_DIR, TELEGRAM_FILE_LIMIT
from core.periods import period_sort_key, fmt_period, is_period_key, is_yearly_key
from core.storage import (
    load_master, save_master, load_state, save_state,
    import_existing_csvs, get_full_period_map,
)
from ai_analyzer import analyze_iok

logger = logging.getLogger(__name__)


async def _run_iok_analysis_for_download(bot, chat_id: int) -> None:
    """Run TENGENOMIKA-style analysis on the most recent IOK periods after Excel download."""
    try:
        df = load_master()
        if df is None or df.empty:
            return
        pm = get_full_period_map(df)
        if not pm:
            return
        keys = list(pm.keys())[-3:]
        labels = [pm[k] for k in keys]

        await bot.send_message(chat_id, "🤖 Генерирую аналитику...")
        analysis = await asyncio.get_event_loop().run_in_executor(
            None, analyze_iok, df, keys, labels
        )
        if analysis:
            try:
                await bot.send_message(chat_id, analysis, parse_mode="Markdown")
            except Exception:
                await bot.send_message(chat_id, analysis)
    except Exception as e:
        logger.warning(f"AI-анализ ИОК (download) не удался: {e}")


# ─────────────────────────────────────────────
# СКАЧАТЬ EXCEL
# ─────────────────────────────────────────────

async def cb_download(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Отображает список доступных Excel файлов."""
    q = update.callback_query
    await q.answer()

    xlsx = sorted(OUTPUT_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    if not xlsx:
        await q.edit_message_text(
            "❌ Excel файлов нет.\nСначала загрузи данные.",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🏠 Меню", callback_data="main_menu")
            ]])
        )
        return

    if len(xlsx) == 1:
        f_path = xlsx[0]
        file_size = f_path.stat().st_size
        if file_size > TELEGRAM_FILE_LIMIT:
            await q.edit_message_text(
                f"⚠️ Файл слишком большой для Telegram ({file_size // 1024 // 1024} МБ).\n"
                f"Путь на сервере: `{f_path}`",
                parse_mode="Markdown",
                reply_markup=InlineKeyboardMarkup([[
                    InlineKeyboardButton("🏠 Меню", callback_data="main_menu")
                ]])
            )
            return
        await q.edit_message_text(f"📤 Отправляю: {f_path.name}...")
        with open(f_path, "rb") as f:
            await ctx.bot.send_document(q.message.chat_id, f, filename=f_path.name)
        asyncio.create_task(_run_iok_analysis_for_download(ctx.bot, q.message.chat_id))
        return

    ctx.user_data["xlsx_files"] = [str(f) for f in xlsx[:8]]
    kb = []
    for i, f in enumerate(xlsx[:8]):
        sz = f.stat().st_size // 1024
        name_short = f.name[:38] if len(f.name) > 38 else f.name
        kb.append([InlineKeyboardButton(f"{name_short} ({sz} КБ)", callback_data=f"dl_{i}")])
    kb.append([InlineKeyboardButton("🏠 Меню", callback_data="main_menu")])

    await q.edit_message_text(
        "📁 Выбери файл для скачивания:",
        reply_markup=InlineKeyboardMarkup(kb)
    )


async def cb_dl_file(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Отправляет конкретный Excel файл."""
    q = update.callback_query
    await q.answer()

    idx = int(q.data.split("_")[1])
    files = ctx.user_data.get("xlsx_files", [])

    if idx >= len(files):
        await q.edit_message_text("❌ Файл не найден")
        return

    from pathlib import Path
    fpath = Path(files[idx])

    if not fpath.exists():
        await q.edit_message_text("❌ Файл не найден на диске")
        return

    file_size = fpath.stat().st_size
    if file_size > TELEGRAM_FILE_LIMIT:
        await q.edit_message_text(
            f"⚠️ Файл слишком большой для Telegram ({file_size // 1024 // 1024} МБ).\n"
            f"Путь на сервере: `{fpath}`",
            parse_mode="Markdown",
        )
        return

    await q.edit_message_text(f"📤 Отправляю {fpath.name}...")
    with open(fpath, "rb") as f:
        await ctx.bot.send_document(q.message.chat_id, f, filename=fpath.name)
    asyncio.create_task(_run_iok_analysis_for_download(ctx.bot, q.message.chat_id))


# ─────────────────────────────────────────────
# ИМПОРТ СУЩЕСТВУЮЩИХ CSV
# ─────────────────────────────────────────────

async def cb_import_csv(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Инструкция по импорту CSV."""
    q = update.callback_query
    await q.answer()
    await q.edit_message_text(
        "📤 *Импорт существующих CSV — один раз*\n\n"
        "Положи в папку `taldau_bot/` два файла:\n"
        "• `ALL_KAZAKHSTAN_FULL.csv`\n"
        "• `MONTHLY_FINAL.csv`\n\n"
        "После этого бот создаст мастер-базу "
        "и будет только дописывать новые месяцы.",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Файлы готовы — импортировать", callback_data="do_import")],
            [InlineKeyboardButton("❌ Отмена", callback_data="main_menu")],
        ])
    )


async def cb_do_import(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Выполняет импорт CSV файлов в мастер-базу."""
    q = update.callback_query
    await q.answer()
    await q.edit_message_text("⏳ Импортирую CSV файлы в базу...")

    def do():
        # Ищем файлы от BASE_DIR (папка проекта), не от текущей рабочей директории
        full_path = BASE_DIR / "ALL_KAZAKHSTAN_FULL.csv"
        monthly_path = BASE_DIR / "MONTHLY_FINAL.csv"
        if not full_path.exists():
            return None, f"❌ Файл ALL_KAZAKHSTAN_FULL.csv не найден в {BASE_DIR}"
        if not monthly_path.exists():
            return None, f"❌ Файл MONTHLY_FINAL.csv не найден в {BASE_DIR}"
        df = import_existing_csvs(str(full_path), str(monthly_path))
        return df, None

    df, err = await asyncio.get_event_loop().run_in_executor(None, do)

    if err:
        await q.edit_message_text(
            err,
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton("🏠 Меню", callback_data="main_menu")
            ]])
        )
        return

    save_master(df)

    # Автоматически определяем и разделяем загруженные периоды
    all_pcols = [c for c in df.columns if is_period_key(c)]
    yearly_loaded = sorted([c for c in all_pcols if is_yearly_key(c)], key=period_sort_key)
    monthly_loaded = sorted([c for c in all_pcols if not is_yearly_key(c)], key=period_sort_key)

    state = load_state()
    state["loaded_periods"] = monthly_loaded
    state["loaded_yearly"] = yearly_loaded
    from datetime import datetime
    state["last_update"] = datetime.now().isoformat()
    save_state(state)

    pm = get_full_period_map(df)
    labels = list(pm.values())
    await q.edit_message_text(
        f"✅ Импорт завершён!\n\n"
        f"📦 Строк: *{len(df):,}*\n"
        f"📅 Периодов: {len(pm)}\n"
        f"  📆 Годовых: {len(yearly_loaded)}\n"
        f"  📅 Месячных: {len(monthly_loaded)}\n"
        f"  Первый: {labels[0] if labels else '—'}\n"
        f"  Последний: {labels[-1] if labels else '—'}\n\n"
        f"Теперь при обновлении будут дописываться только новые данные.",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("📥 Скачать Excel", callback_data="download")],
            [InlineKeyboardButton("🏠 Меню", callback_data="main_menu")],
        ])
    )
