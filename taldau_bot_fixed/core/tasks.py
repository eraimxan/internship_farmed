"""
core/tasks.py — Главный пайплайн: парсинг → обновление мастера → генерация Excel.

Запускается как asyncio Task и шлёт прогресс в Telegram.
"""

import asyncio
import logging
from pathlib import Path

import pandas as pd

from config.settings import (
    DATA_DIR, OUTPUT_DIR, TELEGRAM_FILE_LIMIT,
    PERIOD_ID_MONTHLY, PERIOD_ID_YEARLY,
)
from config.regions import ALL_REGIONS
from core.parser import parse_all_regions
from core.storage import (
    load_master, save_master, merge_new_data,
    mark_periods_loaded, get_full_period_map,
)
from core.excel import build_excel, generate_excel_filename
from ai_analyzer import analyze_iok

logger = logging.getLogger(__name__)


async def run_parse_and_update(
    chat_id: int,
    bot,
    cookie: str,
    period_map: dict[str, str],
    data_type: str = "monthly",
) -> None:
    """
    Основная задача:
      1. Парсит только запрошенные периоды
      2. Дописывает их в мастер (старые данные не трогает)
      3. Строит Excel со ВСЕМИ данными
      4. Отправляет файл в Telegram

    Args:
        chat_id: ID чата для отправки результатов
        bot: экземпляр Telegram Bot
        cookie: строка Cookie для авторизации
        period_map: {period_key: period_label} — только новые/запрошенные периоды
        data_type: "monthly" или "yearly"
    """
    new_cols = list(period_map.keys())
    progress_path = DATA_DIR / "parse_progress.csv"
    total_districts = sum(len(v) for v in ALL_REGIONS.values())

    # Определяем period_id для запроса
    period_id = PERIOD_ID_YEARLY if data_type == "yearly" else PERIOD_ID_MONTHLY

    type_label = "годовых" if data_type == "yearly" else "месячных"

    # Проверяем есть ли незаконченный прогресс
    resume_msg = ""
    if progress_path.exists():
        try:
            import pandas as _pd
            _prev = _pd.read_csv(progress_path, encoding="utf-8-sig", low_memory=False)
            if not _prev.empty and "oblast" in _prev.columns and "region" in _prev.columns:
                done_regions = len(set(zip(_prev["oblast"].astype(str), _prev["region"].astype(str))))
                resume_msg = f"\n🔄 Продолжение: {done_regions}/{total_districts} районов уже пройдено, {len(_prev):,} строк"
            del _prev
        except Exception:
            pass

    await bot.send_message(
        chat_id,
        f"🚀 *Запуск парсинга ({type_label} данных)*\n"
        f"Периодов: {len(period_map)}\n"
        f"Районов: {total_districts}{resume_msg}\n"
        f"⏳ Ожидаемое время: 30–90 минут",
        parse_mode="Markdown",
    )

    # ── Шаг 1: Парсинг ────────────────────────────────────────────
    def do_parse():
        return parse_all_regions(
            cookie=cookie,
            new_cols=new_cols,
            all_regions=ALL_REGIONS,
            period_id=period_id,
            progress_path=progress_path,
        )

    # Глобальный тайм-аут: 3 часа максимум на парсинг
    PARSE_TIMEOUT = 3 * 60 * 60  # секунды
    try:
        all_rows, errors = await asyncio.wait_for(
            asyncio.get_event_loop().run_in_executor(None, do_parse),
            timeout=PARSE_TIMEOUT,
        )
    except asyncio.TimeoutError:
        await bot.send_message(
            chat_id,
            f"⏰ Парсинг превысил лимит времени ({PARSE_TIMEOUT // 3600} ч).\n"
            f"Попробуй загрузить меньше периодов за раз или проверь соединение.",
        )
        return

    if not all_rows:
        await bot.send_message(
            chat_id,
            "❌ Парсинг вернул 0 строк.\n"
            "Проверь Cookie: /setcookie",
        )
        return

    await bot.send_message(
        chat_id,
        f"✅ Парсинг завершён: *{len(all_rows):,}* строк\n"
        f"📊 Обновляю базу данных...",
        parse_mode="Markdown",
    )

    # ── Шаг 2: Обновление мастера ─────────────────────────────────
    def do_merge():
        new_df = pd.DataFrame(all_rows)
        new_df["depth"] = new_df["depth"].fillna(0).astype(int)
        new_df["id"] = pd.to_numeric(
            new_df.get("id", pd.Series(dtype="object")), errors="coerce"
        ).astype("Int64")
        master = load_master()
        updated = merge_new_data(master, new_df, new_cols)
        save_master(updated)
        return updated

    master = await asyncio.get_event_loop().run_in_executor(None, do_merge)

    # Обновляем state.json
    mark_periods_loaded(new_cols, data_type=data_type)

    # Очищаем файл прогресса — парсинг завершён успешно
    try:
        progress_path.unlink(missing_ok=True)
    except Exception:
        pass
    # Также удаляем старый done.txt если остался
    try:
        (DATA_DIR / "parse_done.txt").unlink(missing_ok=True)
    except Exception:
        pass

    await bot.send_message(
        chat_id,
        f"✅ База обновлена: *{len(master):,}* строк\n"
        f"📊 Строю Excel...",
        parse_mode="Markdown",
    )

    # ── Шаг 3: Генерация Excel ────────────────────────────────────
    def do_excel():
        pm = get_full_period_map(master)
        fname = generate_excel_filename(pm)
        fpath = str(OUTPUT_DIR / fname)
        info = build_excel(master, pm, fpath)
        total_rows = sum(v["rows"] for v in info.values())
        return fpath, fname, total_rows

    fpath, fname, total_rows = await asyncio.get_event_loop().run_in_executor(None, do_excel)

    # ── Шаг 4: Отправка файла ─────────────────────────────────────
    file_size = Path(fpath).stat().st_size
    sz_kb = file_size // 1024
    period_labels = list(period_map.values())
    added_str = ", ".join(period_labels[:3]) + ("..." if len(period_labels) > 3 else "")

    caption = (
        f"📊 *Инвестиции в основной капитал — Казахстан*\n"
        f"Добавлено: {added_str}\n"
        f"Всего строк: {total_rows:,} | Размер: {sz_kb} КБ"
    )
    if errors:
        caption += f"\n⚠️ Ошибок при парсинге: {len(errors)}"

    # Проверяем размер файла перед отправкой
    if file_size > TELEGRAM_FILE_LIMIT:
        await bot.send_message(
            chat_id,
            f"⚠️ Файл слишком большой для Telegram ({sz_kb // 1024} МБ).\n"
            f"Файл сохранён на сервере: `{fpath}`\n"
            f"Скачай его вручную с сервера.",
            parse_mode="Markdown",
        )
    else:
        with open(fpath, "rb") as f:
            await bot.send_document(
                chat_id, f,
                filename=fname,
                caption=caption,
                parse_mode="Markdown",
            )

    if errors and len(errors) <= 15:
        err_text = "⚠️ Ошибки при парсинге:\n" + "\n".join(f"• {e}" for e in errors)
        await bot.send_message(chat_id, err_text)
    elif errors:
        await bot.send_message(
            chat_id,
            f"⚠️ Всего ошибок: {len(errors)}. "
            f"Проверь логи для деталей.",
        )
    # ── Шаг 5: AI-анализ ИОК ─────────────────────────────────────
    try:
        await bot.send_message(chat_id, "🤖 Генерирую аналитику...")
        analysis = await asyncio.get_event_loop().run_in_executor(
            None,
            analyze_iok,
            master,
            new_cols,
            list(period_map.values()),
        )
        if analysis:
            try:
                await bot.send_message(chat_id, analysis, parse_mode="Markdown")
            except Exception:
                await bot.send_message(chat_id, analysis)
    except Exception as e:
        logger.warning(f"AI-анализ ИОК не удался: {e}")