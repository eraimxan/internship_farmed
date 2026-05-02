"""
main.py — Точка входа Telegram-бота Taldau.

Запуск:
    BOT_TOKEN=<token> python main.py
    или вставь токен в config/settings.py
"""

import asyncio
import logging
import sys
from pathlib import Path

from telegram import Update, BotCommand
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ── Добавляем корень проекта в sys.path ──────────────────────────
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config.settings import BOT_TOKEN, LOG_FILE
from handlers.menu import cmd_start, cmd_help, cmd_status, cb_main_menu, cb_iok_menu, cb_iok_more, cb_status
from handlers.cookie import cmd_setcookie, cb_set_cookie, handle_cookie_input
from handlers.update import (
    cb_check, cb_update_auto,
    cb_confirm_auto_monthly, cb_confirm_auto_yearly,
    cb_manual_start, cb_sel_year, cb_sel_month, cb_confirm_manual,
    cb_yearly_start, cb_sel_yearly,
    cb_resume, cb_do_resume, cb_reset_progress,
)
from handlers.download import cb_download, cb_dl_file
from handlers.cpi import (
    cb_cpi_menu, cb_cpi_check_upd, cb_cpi_fetch_fresh, cb_cpi_download, cb_cpi_status,
    cb_cpi_do_check, cb_cpi_do_fetch, cb_cpi_do_download, cb_cpi_dl_back,
    cb_cpi_dl_filter, cb_cpi_dl_year, cb_cpi_dl_month, cb_cpi_dl_quarter,
    cb_cpi_sel_period,
)

# ─────────────────────────────────────────────────────────────────
# ЛОГИРОВАНИЕ
# ─────────────────────────────────────────────────────────────────

def setup_logging() -> None:
    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    handlers = [
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ]
    logging.basicConfig(format=fmt, level=logging.INFO, handlers=handlers)
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("telegram").setLevel(logging.WARNING)


logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────
# КОМАНДА /stop
# ─────────────────────────────────────────────────────────────────

async def cmd_stop(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """
    /stop — отменяет все активные фоновые задачи парсинга.
    """
    tasks = [
        t for t in asyncio.all_tasks()
        if t.get_name().startswith("parse_task_")
        and not t.done()
    ]
    if not tasks:
        await update.message.reply_text(
            "ℹ️ Нет активных задач парсинга.\n\n"
            "Если хочешь остановить бот полностью — используй Ctrl+C на сервере."
        )
        return

    for t in tasks:
        t.cancel()

    await update.message.reply_text(
        f"🛑 Отправлен сигнал остановки для {len(tasks)} задач.\n"
        f"Текущий запрос к сайту может ещё завершиться — "
        f"это нормально, следующие запросы уже не пойдут."
    )
    logger.info(f"Пользователь {update.effective_user.id} остановил {len(tasks)} задач")


# ─────────────────────────────────────────────────────────────────
# ТЕКСТОВЫЕ СООБЩЕНИЯ
# ─────────────────────────────────────────────────────────────────

async def handle_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Обрабатывает все текстовые сообщения (в первую очередь — ввод Cookie)."""
    handled = await handle_cookie_input(update, ctx)
    if not handled:
        await update.message.reply_text(
            "Используй /start для главного меню или /help для справки."
        )


# ─────────────────────────────────────────────────────────────────
# РОУТЕР CALLBACK-КНОПОК
# ─────────────────────────────────────────────────────────────────

async def callback_router(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Маршрутизирует все callback_query к нужному хендлеру."""
    data = update.callback_query.data

    routes = {
        "iok_menu":           cb_iok_menu,
        "iok_more":           cb_iok_more,
        "check":                cb_check,
        "status":               cb_status,
        "set_cookie":           cb_set_cookie,
        "update_auto":          cb_update_auto,
        "confirm_auto_monthly": cb_confirm_auto_monthly,
        "confirm_auto_yearly":  cb_confirm_auto_yearly,
        "manual_start":         cb_manual_start,
        "yearly_start":       cb_yearly_start,
        "download":           cb_download,
        "main_menu":          cb_main_menu,
        "confirm_manual":     cb_confirm_manual,
        "resume":             cb_resume,
        "do_resume":          cb_do_resume,
        "reset_progress":     cb_reset_progress,
        "cpi_menu":           cb_cpi_menu,
        "cpi_check_upd":      cb_cpi_check_upd,
        "cpi_fetch_fresh":    cb_cpi_fetch_fresh,
        "cpi_download":       cb_cpi_download,
        "cpi_status":         cb_cpi_status,
        "cpi_dl_back":        cb_cpi_dl_back,
        "cpi_df_all":         cb_cpi_dl_filter,
        "cpi_df_yr":          cb_cpi_dl_filter,
        "cpi_df_mo":          cb_cpi_dl_filter,
        "cpi_df_qr":          cb_cpi_dl_filter,
        "cpi_sel_all":        cb_cpi_sel_period,
    }

    try:
        if data in routes:
            await routes[data](update, ctx)
        elif data.startswith("sel_year_"):
            await cb_sel_year(update, ctx)
        elif data.startswith("sel_month_"):
            await cb_sel_month(update, ctx)
        elif data.startswith("sel_yearly_"):
            await cb_sel_yearly(update, ctx)
        elif data.startswith("dl_"):
            await cb_dl_file(update, ctx)
        elif data.startswith("cpi_cu"):
            await cb_cpi_do_check(update, ctx)
        elif data.startswith("cpi_sel_"):
            await cb_cpi_sel_period(update, ctx)
        elif data.startswith("cpi_nr"):
            await cb_cpi_do_fetch(update, ctx)
        elif data.startswith("cpi_dp"):
            await cb_cpi_do_download(update, ctx)
        elif data.startswith("cpi_yf_") or data.startswith("cpi_ym_") or data.startswith("cpi_yq_"):
            await cb_cpi_dl_year(update, ctx)
        elif data.startswith("cpi_mo_"):
            await cb_cpi_dl_month(update, ctx)
        elif data.startswith("cpi_qu_"):
            await cb_cpi_dl_quarter(update, ctx)
        else:
            logger.warning(f"Неизвестный callback: {data}")
            await update.callback_query.answer("Неизвестное действие")
    except Exception as e:
        if "Message is not modified" not in str(e):
            logger.error(f"Callback error [{data}]: {e}")


# ─────────────────────────────────────────────────────────────────
# ПАТЧ run_parse_and_update ДЛЯ ИМЕНОВАНИЯ ЗАДАЧ (нужен для /stop)
# ─────────────────────────────────────────────────────────────────

import core.tasks as _tasks_module
_original_run = _tasks_module.run_parse_and_update


async def _named_parse_task(chat_id, bot, cookie, period_map, data_type="monthly"):
    """Обёртка которая даёт задаче имя parse_task_* для /stop."""
    task = asyncio.current_task()
    if task:
        task.set_name(f"parse_task_{chat_id}")
    await _original_run(chat_id, bot, cookie, period_map, data_type=data_type)


# Подменяем функцию в модуле tasks
_tasks_module.run_parse_and_update = _named_parse_task


# ─────────────────────────────────────────────────────────────────
# ЗАПУСК
# ─────────────────────────────────────────────────────────────────

async def post_init(app: Application) -> None:
    """Устанавливает команды бота в меню Telegram."""
    await app.bot.set_my_commands([
        BotCommand("start",     "Главное меню"),
        BotCommand("status",    "Статус базы данных"),
        BotCommand("setcookie", "Обновить Cookie сессии"),
        BotCommand("stop",      "Остановить фоновый парсинг"),
        BotCommand("help",      "Справка по командам"),
    ])
    logger.info("Команды бота зарегистрированы")


def main() -> None:
    setup_logging()

    if BOT_TOKEN == "YOUR_BOT_TOKEN_HERE":
        logger.error("Токен бота не задан!")
        print(
            "\n❌ Ошибка: токен бота не задан.\n"
            "Способ 1 (рекомендуется): BOT_TOKEN=<токен> python main.py\n"
            "Способ 2: вставь токен в config/settings.py в переменную BOT_TOKEN\n"
        )
        sys.exit(1)

    logger.info("Запуск Taldau Bot...")

    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .post_init(post_init)
        .build()
    )

    # Команды
    app.add_handler(CommandHandler("start",     cmd_start))
    app.add_handler(CommandHandler("help",      cmd_help))
    app.add_handler(CommandHandler("status",    cmd_status))
    app.add_handler(CommandHandler("setcookie", cmd_setcookie))
    app.add_handler(CommandHandler("stop",      cmd_stop))

    # Callback-кнопки
    app.add_handler(CallbackQueryHandler(callback_router))

    # Текстовые сообщения (не команды)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("✅ Бот запущен. Для остановки нажми Ctrl+C.")
    print("✅ Taldau Bot запущен. Ctrl+C для остановки.")

    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
