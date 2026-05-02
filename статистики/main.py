import asyncio
import html
import traceback
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.client.session.aiohttp import AiohttpSession

from nb_deposits import generate_report
from nb_monetary_base import build_excel_monetary_base
from nb_reserves import build_excel_reserves
from nb_banks_stats import build_excel_banks_stats
from nb_investments_capital import build_excel_investments
from mf_budget_income_expenses import write_summary_workbook
from nb_fdi_report import build_excel_pii
from nb_reer import build_excel_with_chart
from nb_netto_usd import build_usd_excel_with_chart
from nb_fx_rates import build_fx_rates_excel
from nb_bvu_balance import build_excel_bvu_balance_with_chart
from nb_bvu_income_expenses import build_excel_bvu_income_expenses
from mf_budget_dynamics import create_table_in_excel
from mf_budget_income_report import build_excel_income_report
from mf_budget_taxes import build_excel_taxes

import aiohttp
import datetime
import pandas as pd
import psycopg2
from datetime import date

# ---------------- Настройки бота ----------------
#  - тестовый
#  - основной
API_TOKEN = "7581885889:AAE1JZVl6nBl_PyBq8fWNTAN40X9EK40dh4"
PROXY_URL = None  # Set to e.g. "socks5://user:pass@host:port" if Telegram is blocked
session = AiohttpSession(proxy=PROXY_URL) if PROXY_URL else None
bot = Bot(token=API_TOKEN, session=session) if session else Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

API_LIST = {
    "РЭОК": "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=299&date_from=2014-01-01&date_to=2050-01-01&lang=ru",
    "Денежная база": "https://nationalbank.kz/ru/monetarybase/denezhnaya-baza-i-agregaty-shirokoy-denezhnoy-massy/records",
    "Международные резервы": "https://nationalbank.kz/ru/international-reserve-and-asset/mezhdunarodnye-rezervy-i-aktivy-nacionalnogo-fonda-rk/records",
    "ПИИ": "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=310&date_from=2006-01-01&date_to=2050-01-01&lang=ru",
    "Инвестиции в основной капитал": "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=280&date_from=2000-01-01&date_to=2050-01-01&lang=ru",
    "Нетто продажи USD": "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=28&date_from=2015-01-01&date_to=2050-01-01&lang=ru",
    "Статистика по банкам": "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=3&date_from=2014-01-01&date_to=2050-01-01&lang=ru",
    "Депозиты": "https://nationalbank.kz/ru/depositoryorganizationsdeposits/depozity-v-depozitnyh-organizaciyah-/records",
    "Курсы валют": "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=33&date_from=2014-01-01&date_to=2050-01-01&lang=ru",
    "Баланс БВУ": "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=1&date_from=2019-01-01&date_to=2050-01-01&lang=ru",
    "Доходы и расходы БВУ": "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=3&date_from=2019-01-01&date_to=2050-01-01&lang=ru",
}

LAST_DATES_FILE = "last_api_dates.csv"
CHECK_INTERVAL_HOURS = 12
ADMIN_CHAT_ID = 1269409477 # Telegram ID
SUBSCRIBERS_FILE = "subscribers.txt"


# ---------------- FSM ----------------
class DateRange(StatesGroup):
    year = State()
    report_type = State()


YEARS = [str(y) for y in range(2020, 2026)]


# ---------------- ГЛАВНОЕ МЕНЮ ----------------
def main_menu():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🏦 Национальный банк РК", callback_data="cat_nationalbank"),
                InlineKeyboardButton(text="💰 Бюджет", callback_data="cat_budget")
            ],
            [
                InlineKeyboardButton(text="🏛️ Минфин", callback_data="cat_minfin"),
                InlineKeyboardButton(text="📈 БНС АСПР", callback_data="cat_bnss")
            ],
            [
                InlineKeyboardButton(text="🏗️ Миннацэкономики", callback_data="cat_mne"),
                InlineKeyboardButton(text="🌍 ООН (ФАО)", callback_data="cat_un")
            ],
            [
                InlineKeyboardButton(text="📊 Прочие источники", callback_data="cat_other")
            ],
            [
                InlineKeyboardButton(text="🔔 Проверить обновления", callback_data="check_updates_now")
            ],
            [
                InlineKeyboardButton(text="🔔 Подписаться на уведомления", callback_data="subscribe_now")
            ]
        ]
    )


# ---------------- МЕНЮ КАТЕГОРИЙ ----------------
def category_menu(category: str):
    """Возвращает клавиатуру по выбранной категории"""
    if category == "cat_nationalbank":
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="Депозиты в депозитных организациях", callback_data="deposits")],
                [InlineKeyboardButton(text="Денежная база", callback_data="money_base")],
                [InlineKeyboardButton(text="Международные резервы", callback_data="reserves")], 
                [InlineKeyboardButton(text="Кредиты экономике в расширенном определении", callback_data="no_data")],#пустой
                [InlineKeyboardButton(text="Нетто продажи валюты", callback_data="netto_usd")],
                [InlineKeyboardButton(text="Сведения о микрокредитах", callback_data="no_data")], #нет в списке 
                [InlineKeyboardButton(text="РЭОК", callback_data="real_exchange_rate")], 
                [InlineKeyboardButton(text="Чистые продажи", callback_data="no_data")], #нет апишки
                [InlineKeyboardButton(text="Статистика по банкам", callback_data="banks_stats")],
                [InlineKeyboardButton(text="Прямые иностранные инвестиции ", callback_data="foreign_investments")],
                [InlineKeyboardButton(text="Инвестиции в основной капитал", callback_data="investments_capital")],
                [InlineKeyboardButton(text="Реализация золота населению", callback_data="no_data")], #нет в списке
                [InlineKeyboardButton(text="Состояние текущего счета платежного баланса", callback_data="no_data")], #есть в табличной форме
                [InlineKeyboardButton(text="Официальные обменные курсы иностранных валют", callback_data="valuta")],
                [InlineKeyboardButton(text="Баланс БВУ РК", callback_data="bvu_balance")],
                [InlineKeyboardButton(text="Доходы и расходы БВУ РК", callback_data="bvu_income_expenses")],
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")],
            ]
        )
        title = "🏦 <b>Национальный банк Республики Казахстан</b>"
        return title, kb

    elif category == "cat_budget":
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="Доходы исполнение респ бюджета", callback_data="budget_income")],
                [InlineKeyboardButton(text="Доходы налоги Казахстана", callback_data="budget_taxes")],
                [InlineKeyboardButton(text="Динамика 2015–2025 + последний месяц", callback_data="budget_dynamic")],
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")],
            ]
        )
        title = "💰 <b>Бюджет</b>"
        return title, kb

    elif category == "cat_minfin":
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="Доходы исполнение респ бюджета", callback_data="budget_income")],
                [InlineKeyboardButton(text="Доходы налоги Казахстана", callback_data="budget_taxes")],
                [InlineKeyboardButton(text="Динамика 2015–2025 + последний месяц", callback_data="budget_dynamic")],
                [InlineKeyboardButton(text="Государственный долг", callback_data="no_data")],
                [InlineKeyboardButton(text="Показатели НацФонда", callback_data="no_data")],
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")],
            ]
        )
        title = "🏛️ <b>Министерство финансов РК</b>"
        return title, kb

    elif category == "cat_bnss":
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="Уровень рентабельности средних & крупных предприятий", callback_data="no_data")],
                [InlineKeyboardButton(text="Динамика ИФО ВВП", callback_data="no_data")],
                [InlineKeyboardButton(text="Основные показатели рынка труда", callback_data="no_data")],
                [InlineKeyboardButton(text="Продажа и постановка на учет авто", callback_data="no_data")],
                [InlineKeyboardButton(text="Внешняя торговля Казахстана", callback_data="no_data")],
                [InlineKeyboardButton(text="Показатели базовой инфляции", callback_data="no_data")],
                [InlineKeyboardButton(text="Месячная инфляция Казахстана", callback_data="no_data")],
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")],
            ]
        )
        title = "📈 <b>Бюро национальной статистики АСПР</b>"
        return title, kb

    elif category == "cat_mne":
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="Рост экономики Казахстана", callback_data="no_data")],
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")],
            ]
        )
        title = "🏗️ <b>Министерство нацэкономики РК</b>"
        return title, kb

    elif category == "cat_un":
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="Динамика индекса ФАО", callback_data="no_data")],
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")],
            ]
        )
        title = "🌍 <b>ООН (ФАО)</b>"
        return title, kb

    elif category == "cat_other":
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="ОПЕК", callback_data="no_data")],
                [InlineKeyboardButton(text="Цены на жильё и инфляция", callback_data="no_data")],
                [InlineKeyboardButton(text="🔙 Назад", callback_data="back_to_main")],
            ]
        )
        title = "📊 <b>Прочие источники</b>"
        return title, kb

    return "Раздел не найден", InlineKeyboardMarkup()

# ---------------- Вспомогательные функции ----------------
def get_dynamic_period(selected_year: int, records: list):
    """

    Определяет период на основе последней доступной даты в API.
    Если данных меньше 12 месяцев — ограничивает до последней даты.


    """
    available_dates = [
        pd.to_datetime(r.get("reporting_date")).date()
        for r in records if r.get("reporting_date")
    ]
    if not available_dates:
        return None, None

    last_date = max(available_dates)

    # Если пользователь выбрал тот же год, что последняя дата — берем последние 12 мес
    if selected_year == last_date.year:
        end_date = last_date
        start_date = (end_date.replace(day=1) - datetime.timedelta(days=365)).replace(day=1)
    else:
        # Полный календарный год
        start_date = datetime.date(selected_year, 1, 1)
        end_date = datetime.date(selected_year, 12, 31)

    return start_date, end_date


# ---------------- Обработчики ----------------
@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    text = (
        "👋 <b>Добро пожаловать в TENGENOMIKA BOT</b>\n\n"
        "Этот бот формирует актуальные статистические отчёты по данным:\n"
        "🏦 Национального банка РК\n"
        "💰 Министерства финансов РК\n"
        "📊 БНС АСПР и других источников.\n\n"
        "Выберите категорию ниже 👇"
    )
    await message.answer(text, parse_mode="HTML", reply_markup=main_menu())


@dp.callback_query(lambda c: c.data.startswith("cat_"))
async def show_category(callback: types.CallbackQuery):
    title, kb = category_menu(callback.data)
    await callback.message.edit_text(title, parse_mode="HTML", reply_markup=kb)

@dp.callback_query(lambda c: c.data == "back_to_main")
async def back_to_main(callback: types.CallbackQuery):
    await callback.answer()
    try:
        await callback.message.edit_text("📊 Выберите категорию:", reply_markup=main_menu())
    except Exception:
        # если сообщение без текста — просто отправляем новое
        await callback.message.answer("📊 Выберите категорию:", reply_markup=main_menu())


@dp.callback_query(lambda c: c.data == "cancel")
async def cancel_process(callback: types.CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text("🚫 Действие отменено. Выберите показатель заново:", reply_markup=main_menu())


@dp.callback_query(lambda c: c.data == "no_data")
async def no_data_msg(callback: types.CallbackQuery):
    await callback.answer()
    await callback.message.answer("⚙️ Этот показатель пока в разработке. Следите за обновлениями!")


# --- Начало выбора отчёта ---
@dp.callback_query(lambda c: c.data in ["deposits", "money_base", "reserves", "banks_stats", "foreign_investments", "investments_capital", "real_exchange_rate", "netto_usd", "valuta", "bvu_balance", "bvu_income_expenses", "budget_income_expenses"])
async def start_report(callback: types.CallbackQuery, state: FSMContext):
    await callback.answer()
    report_names = {
        "deposits": "Депозиты в депозитных организациях",
        "money_base": "Денежная база",
        "reserves": "Международные резервы",
        "banks_stats": "Статистика по банкам",
        "investments_capital": "Инвестиции в основной капитал",
        "foreign_investments": "Прямые иностранные инвестиции",
        "real_exchange_rate": "РЭОК",
        "netto_usd": "Нетто продажи валюты",
        "valuta": "Официальные обменные курсы иностранных валют",
        "bvu_balance": "Баланс БВУ РК",
        "bvu_income_expenses": "Доходы и расходы БВУ РК",
    }
    report_type = callback.data
    chosen_report = report_names[report_type]

    await callback.message.edit_text(f"💡 Вы выбрали показатель: <b>{chosen_report}</b>", parse_mode="HTML")

    # --- Определяем последнюю дату обновления ---
    last_date_text = "неизвестна"
    try:
        url = None
        if report_type == "money_base":
            url = "https://nationalbank.kz/ru/monetarybase/denezhnaya-baza-i-agregaty-shirokoy-denezhnoy-massy/records"
        elif report_type == "reserves":
            url = "https://nationalbank.kz/ru/international-reserve-and-asset/mezhdunarodnye-rezervy-i-aktivy-nacionalnogo-fonda-rk/records"
        elif report_type == "banks_stats":
            url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=3&date_from=2014-01-01&date_to=2050-01-01&lang=ru"
        elif report_type == "foreign_investments":
            url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=310&date_from=2006-01-01&date_to=2050-01-01&lang=ru"
        elif report_type == "deposits":
            url = "https://nationalbank.kz/ru/depositoryorganizationsdeposits/depozity-v-depozitnyh-organizaciyah-/records"
        elif report_type == "investments_capital":
            url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=280&date_from=2000-01-01&date_to=2050-01-01&lang=ru"
        elif report_type == "real_exchange_rate":
            url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=299&date_from=2014-01-01&date_to=2050-01-01&lang=ru"
        elif report_type == "netto_usd":
            url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=28&date_from=2015-01-01&date_to=2050-01-01&lang=ru"
        elif report_type == "valuta":
            url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=33&date_from=2014-01-01&date_to=2050-01-01&lang=ru"
        elif report_type == "bvu_balance":
            url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=1&date_from=2014-01-01&date_to=2050-01-01&lang=ru"
        elif report_type == "bvu_income_expenses":
            url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=3&date_from=2025-01-01&date_to=2050-01-01&lang=ru"

        if url:
            async with aiohttp.ClientSession() as session:
                async with session.get(url, headers={"Accept": "application/json"}) as resp:
                    data_json = await resp.json(content_type=None)
                    if isinstance(data_json, dict):
                        records = data_json.get("data", []) or data_json.get("records", [])
                    elif isinstance(data_json, list):
                        records = data_json
                    else:
                        records = []

            available_dates = []
            for r in records:
                date_val = (
                    r.get("reporting_date")
                    or r.get("reportDate")
                    or r.get("repDate")
                    or r.get("date")
                    or r.get("reportdate")
                )
                if date_val:
                    try:
                        available_dates.append(pd.to_datetime(date_val).date())
                    except Exception:
                        continue

            if available_dates:
                last_date = max(available_dates)
                month_names = {
                    "January": "январь", "February": "февраль", "March": "март", "April": "апрель",
                    "May": "май", "June": "июнь", "July": "июль", "August": "август",
                    "September": "сентябрь", "October": "октябрь", "November": "ноябрь", "December": "декабрь"
                }
                month_rus = month_names.get(last_date.strftime("%B"), last_date.strftime("%B"))
                last_date_text = f"{month_rus} {last_date.year} г."

    except Exception as e:
        print("⚠️ Ошибка при проверке актуальности:", e)

    # --- FSM: сохраняем тип отчёта ---
    await state.set_state(DateRange.year)
    await state.update_data(report_type=report_type)

    # --- Клавиатура годов ---
    kb = InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text=y, callback_data=y)] for y in YEARS] +
                        [[InlineKeyboardButton(text="❌ Отмена", callback_data="cancel")]]
    )

    # --- Выводим актуальность и выбор года ---
    await callback.message.answer(
        f"📅 Данные актуальны по состоянию на: <b>{last_date_text}</b>\n\nВыберите год:",
        parse_mode="HTML",
        reply_markup=kb
    )

# --- Динамика за последний год по ОИБ ---
@dp.callback_query(lambda c: c.data == "budget_dynamic")
async def handle_budget_dynamic(callback: types.CallbackQuery):
    await callback.answer()
    await callback.message.answer("⏳ Формирую отчёт 'Динамика 2015–2025 + последний месяц'...")

    try:
        file_name = "Динамика_2015_2025_последний_месяц.xlsx"
        await asyncio.get_event_loop().run_in_executor(
            None, lambda: create_table_in_excel(None, file_name, oblast_name="")
        )

        file = FSInputFile(file_name)
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="🏠 Главное меню", callback_data="back_to_main")]
            ]
        )

        await callback.message.answer_document(
            file,
            caption="✅ Отчёт готов: 2015–2025 по годам + последний месяц из базы",
            reply_markup=kb
        )

    except Exception:
        tb = html.escape(traceback.format_exc())
        tb = tb[-3000:] if len(tb) > 3000 else tb
        await callback.message.answer(
            f"⚠️ Ошибка при формировании отчёта:\n<code>{tb}</code>",
            parse_mode="HTML"
        )


# --- Доходы исполнение республиканского бюджета ---
@dp.callback_query(lambda c: c.data == "budget_income")
async def handle_budget_income(callback: types.CallbackQuery):
    await callback.answer()
    await callback.message.answer("⏳ Формирую отчёт 'Доходы исполнение республиканского бюджета'...")

    try:
        file_name = "Доходы_исполнение_респ_бюджета.xlsx"
        await asyncio.get_event_loop().run_in_executor(
            None, build_excel_income_report, file_name
        )

        file = FSInputFile(file_name)
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="🏠 Главное меню", callback_data="back_to_main")]
            ]
        )
        await callback.message.answer_document(
            file,
            caption="✅ Отчёт 'Исполнение республиканского бюджета — Доходы' готов.",
            reply_markup=kb
        )

    except Exception as e:
        err = html.escape(traceback.format_exc())
        await callback.message.answer(
            f"⚠️ Ошибка:\n<pre>{err}</pre>",
            parse_mode="HTML"
        )


# --- Доходы налоги Казахстана ---
@dp.callback_query(lambda c: c.data == "budget_taxes")
async def handle_budget_taxes(callback: types.CallbackQuery):
    await callback.answer()
    await callback.message.answer("⏳ Формирую отчёт 'Доходы налоги Казахстана'...")

    try:
        file_name = "Доходы_налоги_Казахстана.xlsx"
        await asyncio.get_event_loop().run_in_executor(
            None, build_excel_taxes, file_name
        )

        file = FSInputFile(file_name)
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="🏠 Главное меню", callback_data="back_to_main")]
            ]
        )
        await callback.message.answer_document(
            file,
            caption="✅ Отчёт 'Налоговые поступления Казахстана' готов.",
            reply_markup=kb
        )

    except Exception as e:
        err = html.escape(traceback.format_exc())
        await callback.message.answer(
            f"⚠️ Ошибка:\n<pre>{err}</pre>",
            parse_mode="HTML"
        )


# --- Формирование отчёта ---
@dp.callback_query(DateRange.year)
async def choose_year(callback: types.CallbackQuery, state: FSMContext):
    if callback.data == "cancel":
        await cancel_process(callback, state)
        return

    year = int(callback.data)
    data = await state.get_data()
    report_type = data["report_type"]

    await callback.message.edit_text("⏳ <b>Формирую отчёт...</b>\nПожалуйста, подождите ⌛", parse_mode="HTML")


    file_path = None

    # ------------------ Депозиты ------------------
    if report_type == "deposits":
        # тут generate_report уже сам подбирает данные, но добавим фиксацию периода для имени файла
        end_date = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
        start_date = (end_date.replace(day=1) - datetime.timedelta(days=365)).replace(day=1)
        caption_period = f"{start_date.strftime('%m.%Y')}–{end_date.strftime('%m.%Y')}"
        file_name = f"Депозиты_{caption_period}.xlsx"
        file_path = await generate_report(start_date.year, start_date.month, end_date.year, end_date.month, file_name)

    # ------------------ Денежная база ------------------
    elif report_type == "money_base":
        url = "https://nationalbank.kz/ru/monetarybase/denezhnaya-baza-i-agregaty-shirokoy-denezhnoy-massy/records"
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers={"Accept": "application/json"}) as resp:
                data_json = await resp.json()
                if isinstance(data_json, dict):
                    records = data_json.get("data", [])
                elif isinstance(data_json, list):
                    records = data_json
                else:
                    records = []

                

        start_date, end_date = get_dynamic_period(year, records)
        if not start_date:
            await callback.message.answer("⚠️ Нет данных по денежной базе.")
            await state.clear()
            return

        caption_period = f"{start_date.strftime('%m.%Y')}–{end_date.strftime('%m.%Y')}"
        filtered = [r for r in records if start_date <= pd.to_datetime(r["reporting_date"]).date() <= end_date]

        if filtered:
            file_path = build_excel_monetary_base(filtered, f"Денежная_база_{caption_period}.xlsx")

    # ------------------ Международные резервы ------------------
    elif report_type == "reserves":
        url = "https://nationalbank.kz/ru/international-reserve-and-asset/mezhdunarodnye-rezervy-i-aktivy-nacionalnogo-fonda-rk/records"
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers={"Accept": "application/json"}) as resp:
                data_json = await resp.json()
                if isinstance(data_json, dict):
                    records = data_json.get("data", [])
                elif isinstance(data_json, list):
                    records = data_json
                else:
                    records = []


        start_date, end_date = get_dynamic_period(year, records)
        if not start_date:
            await callback.message.answer("⚠️ Нет данных по международным резервам.")
            await state.clear()
            return

        caption_period = f"{start_date.strftime('%m.%Y')}–{end_date.strftime('%m.%Y')}"
        filtered = [r for r in records if start_date <= pd.to_datetime(r["reporting_date"]).date() <= end_date]

        if filtered:
            file_path = build_excel_reserves(filtered, f"Международные_резервы_{caption_period}.xlsx")

    # ------------------ Статистика по банкам ------------------
    elif report_type == "banks_stats":
        url = "https://data.nationalbank.kz/api/report/stat-data-locale?form_id=3&date_from=2014-01-01&date_to=2050-01-01&lang=ru"
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers={"Accept": "application/json"}) as resp:
                data_json = await resp.json()
                records = data_json.get("data", []) if isinstance(data_json, dict) else data_json
        filtered = [r for r in records if str(year) in r.get("reportDate", "")]
        if filtered:
            file_path = build_excel_banks_stats(filtered, f"Статистика_по_банкам_{year}.xlsx")
            caption_period = str(year)

    # ------------------ Инвестиции в основной капитал ------------------
    elif report_type == "investments_capital":
        try:
            file_name = f"Инвестиции_в_основной_капитал_{year}.xlsx"
            build_excel_investments(file_name)
            file_path = file_name
            caption_period = str(year)
        except Exception as e:
            print("⚠️ Ошибка при формировании инвестиций в основной капитал:", e)
            file_path = None
    
    # ------------------ Прямые иностранные инвестиции ------------------
    elif report_type == "foreign_investments":
        try:
            file_name = f"Прямые_иностранные_инвестиции_{year}.xlsx"
            build_excel_pii(file_name)
            file_path = file_name
            caption_period = str(year)
        except Exception as e:
            print("⚠️ Ошибка при формировании ПИИ:", e)
            file_path = None

    # ------------------ РЭОК ------------------
    elif report_type == "real_exchange_rate":
        try:
            file_name = f"РЭОК_месячные_{year}.xlsx"
            build_excel_with_chart(file_name)
            file_path = file_name
            caption_period = str(year)
        except Exception as e:
            print("⚠️ Ошибка при формировании РЭОК:", e)
            file_path = None 

    # ------------------ Нетто продажи USD ------------------
    elif report_type == "netto_usd":
        try:
            file_name = f"Нетто_продажи_USD_{year}.xlsx"
            build_usd_excel_with_chart(file_name, year)
            file_path = file_name
            caption_period = str(year)
        except Exception as e:
            print("⚠️ Ошибка при формировании нетто продаж USD:", e)
            file_path = None

    # ------------------ Валюта ------------------
    elif report_type == "valuta":
        try:
            file_name = f"Курсы_валют_{year}.xlsx"
            build_fx_rates_excel(year, file_name)
            file_path = file_name
            caption_period = str(year)
        except Exception as e:
            print("⚠️ Ошибка при формировании курсов валют:", e)
            file_path = None

    # ------------------ Баланс БВУ ------------------
    elif report_type == "bvu_balance":
        try:
            file_name = f"Баланс_БВУ_{year}.xlsx"
            build_excel_bvu_balance_with_chart(file_name, year)
            file_path = file_name
            caption_period = str(year)
        except Exception as e:
            print("⚠️ Ошибка при формировании баланса БВУ:", e)
            file_path = None
        
    # ------------------ Доходы и расходы БВУ ------------------
    elif report_type == "bvu_income_expenses":
        try:
            file_name = f"Доходы_и_расходы_БВУ_{year}.xlsx"
            build_excel_bvu_income_expenses(file_name, year)
            file_path = file_name
            caption_period = str(year)
        except Exception as e:
            print("⚠️ Ошибка при формировании доходов и расходов БВУ:", e)
            file_path = None

    # ------------------ Результат ------------------
    if file_path:
        file = FSInputFile(file_path)

        # создаём клавиатуру
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="🏠 Главное меню", callback_data="back_to_main")]
            ]
        )

        # отправляем документ с кнопкой
        await callback.message.answer_document(
            file,
            caption=f"✅ Отчёт готов за {caption_period}",
            reply_markup=kb
        )
    else:
        await callback.message.answer("⚠️ Нет данных за выбранный период.")

    await state.clear()



@dp.callback_query(lambda c: c.data == "budget_income_expenses")
async def handle_budget_income_expenses(callback: types.CallbackQuery):
    await callback.answer()
    await callback.message.answer("⏳ Формирую отчёт 'Доходы и затраты по регионам с 2023 года'...")

    try:
        # Подключаемся к базе PostgreSQL
        conn = psycopg2.connect(
            host="56.228.80.80",
            dbname="emfreportbpr",
            user="postgres",
            password="159753NBER",
            port="5432"
        )

        # Формируем Excel
        file_name = "Доходы_и_затраты_по_регионам.xlsx"
        write_summary_workbook(conn, file_name)
        conn.close()

        file = FSInputFile(file_name)
        kb = InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="🏠 Главное меню", callback_data="back_to_main")]
            ]
        )
        await callback.message.answer_document(
            file,
            caption="✅ Отчёт 'Доходы и затраты по регионам с 2023 года' готов.",
            reply_markup=kb
        )


    except Exception as e:
        await callback.message.answer(f"⚠️ Ошибка при формировании отчёта:\n<code>{e}</code>", parse_mode="HTML")

# ---------------- Удведомления ----------------

@dp.callback_query(lambda c: c.data == "subscribe_now")
async def handle_subscribe(callback: types.CallbackQuery):
    if save_subscriber(callback.message.chat.id):
        await callback.answer("Подписка активирована ✅", show_alert=True)
    else:
        await callback.answer("Вы уже подписаны 🔔", show_alert=True)


@dp.message(Command("subscribe"))
async def subscribe_updates(message: types.Message):
    if save_subscriber(message.chat.id):
        await message.answer("🔔 Вы подписались на уведомления о новых данных НБ РК.")
    else:
        await message.answer("✅ Вы уже подписаны на уведомления.")


def read_subscribers():
    try:
        with open(SUBSCRIBERS_FILE, "r") as f:
            return [int(line.strip()) for line in f.readlines()]
    except FileNotFoundError:
        return []

def save_subscriber(chat_id):
    subs = read_subscribers()
    if chat_id not in subs:
        subs.append(chat_id)
        with open(SUBSCRIBERS_FILE, "w") as f:
            f.writelines([str(i) + "\n" for i in subs])
        return True
    return False


async def fetch_latest_date(url):
    """Возвращает самую свежую дату из API"""
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers={"Accept": "application/json"}) as resp:
            data = await resp.json(content_type=None)

    if isinstance(data, dict):
        data = data.get("data") or data.get("records") or []

    dates = []
    for item in data:
        d = (
            item.get("reportDate") or item.get("reporting_date")
            or item.get("repDate") or item.get("date") or item.get("reportdate")
        )
        if d:
            try:
                dates.append(pd.to_datetime(d).date())
            except:
                pass
    return max(dates) if dates else None


def read_last_dates():
    try:
        df = pd.read_csv(LAST_DATES_FILE, index_col="name", parse_dates=["last_date"])
        return df["last_date"].to_dict()
    except FileNotFoundError:
        return {}


def save_last_dates(dates_dict):
    df = pd.DataFrame(list(dates_dict.items()), columns=["name", "last_date"])
    df.to_csv(LAST_DATES_FILE, index=False)


async def check_api_updates():
    """Основная функция проверки всех API"""
    while True:
        print(f"\n=== Проверка обновлений {date.today()} ===")
        saved_dates = read_last_dates()
        new_dates = {}
        updates = []
        no_updates = []

        for name, url in API_LIST.items():
            try:
                latest_date = await fetch_latest_date(url)
                if not latest_date:
                    continue
                old_date = saved_dates.get(name)
                new_dates[name] = latest_date

                MONTHS_RU = {
                    1: "январь", 2: "февраль", 3: "март", 4: "апрель",
                    5: "май", 6: "июнь", 7: "июль", 8: "август",
                    9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"
                }

                month_ru = MONTHS_RU.get(latest_date.month, latest_date.strftime("%B"))
                if old_date is None or latest_date > pd.to_datetime(old_date).date():
                    updates.append(f"{name} — {month_ru} {latest_date.year}")
                else:
                    no_updates.append(f"{name} — без изменений")

            except Exception as e:
                print(f"⚠️ Ошибка при {name}: {e}")

        if updates:
            msg = "📊 <b>Обновления данных НБ РК:</b>\n\n"
            msg += "\n".join(f"✅ {u}" for u in updates)
            if no_updates:
                msg += "\n\n⚪ Без изменений:\n" + "\n".join(f"▫️ {n}" for n in no_updates)
            subscribers = read_subscribers()
            for user_id in subscribers:
                try:
                    await bot.send_message(user_id, msg, parse_mode="HTML")
                except Exception as e:
                    print(f"⚠️ Ошибка при отправке {user_id}: {e}")

            print("✅ Отправлено уведомление в Telegram")
        else:
            print("ℹ️ Новых данных нет")

        save_last_dates(new_dates)
        await asyncio.sleep(CHECK_INTERVAL_HOURS * 3600)

# ---------------- Ручная проверка ----------------
@dp.callback_query(lambda c: c.data == "check_updates_now")
async def manual_check_api(callback: types.CallbackQuery):
    await callback.answer("⏳ Проверяю обновления...", show_alert=False)

    saved_dates = read_last_dates()
    new_dates = {}
    updates = []
    no_updates = []

    for name, url in API_LIST.items():
        try:
            latest_date = await fetch_latest_date(url)
            if not latest_date:
                continue
            old_date = saved_dates.get(name)
            new_dates[name] = latest_date

            MONTHS_RU = {
                1: "январь", 2: "февраль", 3: "март", 4: "апрель",
                5: "май", 6: "июнь", 7: "июль", 8: "август",
                9: "сентябрь", 10: "октябрь", 11: "ноябрь", 12: "декабрь"
            }

            month_ru = MONTHS_RU.get(latest_date.month, latest_date.strftime("%B"))
            if old_date is None or latest_date > pd.to_datetime(old_date).date():
                updates.append(f"✅ {name} — {month_ru} {latest_date.year}")
            else:
                no_updates.append(f"▫️ {name} — без изменений")

        except Exception as e:
            print(f"⚠️ Ошибка при {name}: {e}")

    msg = ""
    if updates:
        msg += "📊 <b>Обновления данных НБ РК:</b>\n\n" + "\n".join(updates)
    else:
        msg += "ℹ️ <b>Новых данных не найдено</b>\n"

    if no_updates:
        msg += "\n\n⚪ Без изменений:\n" + "\n".join(no_updates)

    save_last_dates(new_dates)
    await callback.message.answer(msg, parse_mode="HTML")

@dp.message(Command("id"))
async def get_id(message: types.Message):
    await message.answer(f"🆔 Ваш chat_id: <code>{message.chat.id}</code>", parse_mode="HTML")

# ---------------- Запуск ----------------
async def main():
    # Запускаем проверку обновлений в фоне
    asyncio.create_task(check_api_updates())
    # Запускаем бота
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
