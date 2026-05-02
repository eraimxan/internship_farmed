"""
handlers/cookie.py — Управление Cookie сессии.
"""

import asyncio
import logging

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from core.storage import save_cookie, load_state, save_state
from core.parser import check_session_both

logger = logging.getLogger(__name__)

COOKIE_INSTRUCTIONS = (
    "🔑 *Обновление Cookie*\n\n"
    "1. Открой: https://taldau.stat.gov.kz/ru/NewIndex/GetIndex/701830\n"
    "2. F12 → вкладка *Network* → обнови страницу\n"
    "3. Найди запрос `GetIndexTreeData` → Headers → скопируй `Cookie:`\n\n"
    "Отправь Cookie следующим сообщением 👇"
)


async def cmd_setcookie(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Команда /setcookie."""
    await update.message.reply_text(COOKIE_INSTRUCTIONS, parse_mode="Markdown")
    ctx.user_data["waiting_cookie"] = True


async def cb_set_cookie(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    """Callback — запрос обновления cookie через меню."""
    q = update.callback_query
    await q.answer()
    await q.edit_message_text(COOKIE_INSTRUCTIONS, parse_mode="Markdown")
    ctx.user_data["waiting_cookie"] = True


async def handle_cookie_input(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> bool:
    """
    Обрабатывает текстовый ввод Cookie.
    Возвращает True если сообщение было обработано как Cookie.
    """
    if not ctx.user_data.get("waiting_cookie"):
        return False

    ctx.user_data["waiting_cookie"] = False
    cookie = update.message.text.strip()

    if len(cookie) < 20:
        await update.message.reply_text(
            "❌ Похоже это не Cookie. Cookie обычно длинная строка вида:\n"
            "`ASP.NET_SessionId=...; .ASPXAUTH=...`\n\n"
            "Попробуй снова: /setcookie",
            parse_mode="Markdown",
        )
        return True

    await update.message.reply_text("🔄 Проверяю сессию...")

    ok, monthly, yearly = await asyncio.get_event_loop().run_in_executor(
        None, check_session_both, cookie
    )

    if ok:
        save_cookie(cookie)
        state = load_state()
        state["available_periods"] = monthly
        state["available_yearly"] = yearly
        save_state(state)
        await update.message.reply_text(
            f"✅ Cookie сохранён! Сессия активна.\n"
            f"📅 Месячных периодов на сайте: *{len(monthly)}*\n"
            f"📆 Годовых периодов на сайте: *{len(yearly)}*\n\n"
            f"Используй /start чтобы начать.",
            parse_mode="Markdown",
        )
    else:
        save_cookie(cookie)
        await update.message.reply_text(
            "⚠️ Cookie сохранён, но сессия не прошла проверку.\n"
            "Возможно сайт временно недоступен или Cookie устарел.\n"
            "Попробуй получить новый Cookie.",
        )

    return True
