"""
core/cpi_tasks.py — Async pipeline для загрузки ИПЦ: fetch → Excel → Telegram.
"""

import asyncio
import importlib
import json
import logging
import re
import sys
from pathlib import Path

import pandas as pd
from ai_analyzer import analyze_cpi

# cpi_scripts/ не является пакетом, добавляем в sys.path
_CPI_DIR = Path(__file__).resolve().parent.parent / "cpi_scripts"
if str(_CPI_DIR) not in sys.path:
    sys.path.insert(0, str(_CPI_DIR))

from config.settings import TELEGRAM_FILE_LIMIT

_CPI_OUT = _CPI_DIR / "output"

PERIOD_LABELS = {
    "4": "Месяц",
    "5": "Квартал",
    "9": "Месяц с накоплением",
}

MONTH_NAMES = {
    1: "Январь", 2: "Февраль", 3: "Март",    4: "Апрель",
    5: "Май",    6: "Июнь",    7: "Июль",    8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}

QUARTER_MONTHS = {
    1: ["01", "02", "03"],
    2: ["04", "05", "06"],
    3: ["07", "08", "09"],
    4: ["10", "11", "12"],
}

logger = logging.getLogger(__name__)


_PERIOD_FILE_NAMES = {
    "4": "monthly",
    "5": "quarterly",
    "9": "cumulative",
}

# Legacy bare names (cpi_data.csv) — quarterly only
_OLD_PATHS = {
    "csv":        _CPI_OUT / "cpi_data.csv",
    "checkpoint": _CPI_OUT / "cpi_checkpoint.csv",
    "params":     _CPI_OUT / "cpi_params.json",
    "excel":      _CPI_OUT / "CPI_Kazakhstan.xlsx",
}

# Numeric-suffix names from previous version (cpi_data_5.csv etc.)
def _numeric_paths(period_id: str) -> dict:
    return {
        "csv":        _CPI_OUT / f"cpi_data_{period_id}.csv",
        "checkpoint": _CPI_OUT / f"cpi_checkpoint_{period_id}.csv",
        "params":     _CPI_OUT / f"cpi_params_{period_id}.json",
        "excel":      _CPI_OUT / f"CPI_Kazakhstan_{period_id}.xlsx",
    }


def _cpi_paths(period_id: str) -> dict:
    """Returns paths using human-readable names (cpi_data_monthly.csv, etc.)."""
    name = _PERIOD_FILE_NAMES.get(period_id, period_id)
    return {
        "csv":        _CPI_OUT / f"cpi_data_{name}.csv",
        "checkpoint": _CPI_OUT / f"cpi_checkpoint_{name}.csv",
        "params":     _CPI_OUT / f"cpi_params_{name}.json",
        "excel":      _CPI_OUT / f"CPI_Kazakhstan_{name}.xlsx",
    }


def _existing(period_id: str, key: str) -> Path | None:
    """
    Looks for a file in priority order:
      1. Human-readable name  (cpi_data_monthly.csv)
      2. Numeric-suffix name  (cpi_data_4.csv)
      3. Legacy bare name     (cpi_data.csv) — quarterly only
    Used for reading only (status, download).
    """
    p = _cpi_paths(period_id)[key]
    if p.exists():
        return p
    np = _numeric_paths(period_id)[key]
    if np.exists():
        return np
    if period_id == "5":
        old = _OLD_PATHS.get(key)
        if old and old.exists():
            return old
    return None


def _cpi_module(name: str):
    """Импортирует модуль из cpi_scripts/ с кэшированием в sys.modules."""
    key = f"_cpi_{name}"
    if key not in sys.modules:
        sys.modules[key] = importlib.import_module(name)
    return sys.modules[key]


async def run_cpi_pipeline(
    chat_id: int,
    bot,
    cookie: str,
    period_id: str = "5",
    no_resume: bool = False,
    analysis_period_keys: set | None = None,
    filter_type: str | None = None,
    filter_year: str | None = None,
    filter_month: str | None = None,
    filter_quarter: int | None = None,
) -> None:
    """
    Основная задача ИПЦ:
      1. Загружает данные с taldau (21 регион × 2 типа сравнения)
      2. Строит Excel (полный или отфильтрованный по filter_type/year/month/quarter)
      3. Отправляет файл в Telegram
      4. AI-анализ по выбранным периодам

    filter_type: None/"all"/"year"/"month"/"quarter" — фильтр для Excel и AI.
    analysis_period_keys: явные ключи для AI (приоритет над filter_type).
    """
    paths = _cpi_paths(period_id)
    period_label = PERIOD_LABELS.get(period_id, f"период {period_id}")

    filter_desc = ""
    if filter_type and filter_type != "all":
        if filter_type == "month" and filter_month and filter_year:
            filter_desc = f" — {MONTH_NAMES.get(int(filter_month), filter_month)} {filter_year}"
        elif filter_type == "year" and filter_year:
            filter_desc = f" — {filter_year}"
        elif filter_type == "quarter" and filter_quarter and filter_year:
            filter_desc = f" — Q{filter_quarter} {filter_year}"

    await bot.send_message(
        chat_id,
        f"*📈 Запуск загрузки ИПЦ — {period_label}{filter_desc}*\n\n"
        "Регионов: 21\n"
        "Типов сравнения: 2 (г/г и к пред. периоду)\n"
        "⏳ Ожидаемое время: 5–15 минут",
        parse_mode="Markdown",
    )

    # ── Шаг 1: Fetch ──────────────────────────────────────────────
    def do_fetch() -> int:
        cpi_fetch = _cpi_module("fetch")

        if no_resume and paths["checkpoint"].exists():
            paths["checkpoint"].unlink()
            logger.info("CPI: удалён checkpoint (no_resume=True)")

        session = cpi_fetch.make_session(cookie)
        params  = dict(cpi_fetch.KNOWN_PARAMS)
        params["period_id"] = period_id

        if not cpi_fetch.check_connection(session, params):
            raise RuntimeError("Cookie недействителен или сайт недоступен")

        period_name_map, date_list, name_list = cpi_fetch.get_period_names(session, params)

        df = cpi_fetch.fetch_all(
            session,
            params,
            checkpoint_path=paths["checkpoint"],
            cookie=cookie,
        )

        if df.empty:
            raise RuntimeError("Fetch вернул 0 строк")

        df.to_csv(paths["csv"], index=False, encoding="utf-8-sig")

        params_out = {
            **params,
            "date_list":        date_list,
            "period_name_list": name_list,
            "period_name_map":  period_name_map,
        }
        paths["params"].write_text(
            json.dumps(params_out, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return len(df)

    try:
        row_count = await asyncio.wait_for(
            asyncio.get_event_loop().run_in_executor(None, do_fetch),
            timeout=2 * 60 * 60,
        )
    except asyncio.TimeoutError:
        await bot.send_message(chat_id, "⏰ Загрузка ИПЦ превысила лимит (2 ч). Попробуй снова.")
        return
    except Exception as e:
        await bot.send_message(chat_id, f"❌ Ошибка при загрузке ИПЦ:\n`{e}`", parse_mode="Markdown")
        return

    excel_suffix = filter_desc.lstrip(" —").replace(" ", "_") if filter_desc else ""
    await bot.send_message(
        chat_id,
        f"✅ Данные загружены: *{row_count:,}* строк\n"
        f"📊 Строю Excel{(' — ' + filter_desc.lstrip(' —')) if filter_desc else ''}...",
        parse_mode="Markdown",
    )

    # ── Шаг 2: Excel ──────────────────────────────────────────────
    excel_path  = paths["excel"]
    excel_label = "все периоды"

    if filter_type and filter_type != "all":
        def do_excel():
            return build_filtered_excel(
                period_id, filter_type,
                year=filter_year, month=filter_month, quarter=filter_quarter,
            )
        try:
            excel_path, excel_label = await asyncio.get_event_loop().run_in_executor(None, do_excel)
        except Exception as e:
            await bot.send_message(chat_id, f"❌ Ошибка при построении Excel:\n`{e}`", parse_mode="Markdown")
            return
    else:
        def do_excel_full():
            cpi_excel = _cpi_module("build_excel")
            df = pd.read_csv(paths["csv"], encoding="utf-8-sig", low_memory=False)
            cpi_excel.build_excel(
                df,
                _existing(period_id, "params"),
                paths["excel"],
            )
        try:
            await asyncio.get_event_loop().run_in_executor(None, do_excel_full)
        except Exception as e:
            await bot.send_message(chat_id, f"❌ Ошибка при построении Excel:\n`{e}`", parse_mode="Markdown")
            return

    # ── Шаг 3: Отправка файла ─────────────────────────────────────
    if not excel_path.exists():
        await bot.send_message(chat_id, "❌ Excel файл не создан.")
        return

    file_size = excel_path.stat().st_size
    sz_kb     = file_size // 1024
    caption   = (
        f"📈 *ИПЦ — {period_label}*\n"
        f"Период: {excel_label}\n"
        f"Строк: {row_count:,} | Размер: {sz_kb} КБ"
    )

    if file_size > TELEGRAM_FILE_LIMIT:
        await bot.send_message(
            chat_id,
            f"⚠️ Файл слишком большой ({sz_kb // 1024} МБ) для Telegram.\n"
            f"Сохранён на сервере: `{excel_path}`",
            parse_mode="Markdown",
        )
    else:
        with open(excel_path, "rb") as f:
            await bot.send_document(
                chat_id, f,
                filename=excel_path.name,
                caption=caption,
                parse_mode="Markdown",
            )

    try:
        paths["checkpoint"].unlink(missing_ok=True)
    except Exception:
        pass

    # ── Шаг 4: AI-анализ ИПЦ ─────────────────────────────────────
    try:
        await bot.send_message(chat_id, "🤖 Генерирую аналитику...")
        df_for_analysis = pd.read_csv(paths["csv"], encoding="utf-8-sig", low_memory=False)

        from ai_analyzer import _period_sort_key

        all_period_cols = sorted(
            [c for c in df_for_analysis.columns if re.match(r"^y\d{6}$", c)],
            key=_period_sort_key,
        )

        # Derive analysis keys from filter params if not explicitly provided
        if not analysis_period_keys and filter_type and filter_type != "all":
            if filter_type == "month" and filter_month and filter_year:
                analysis_period_keys = {f"y{filter_month}{filter_year}"}
            elif filter_type == "year" and filter_year:
                analysis_period_keys = {c for c in all_period_cols if c[3:] == filter_year}
            elif filter_type == "quarter" and filter_quarter and filter_year:
                q_months = QUARTER_MONTHS[filter_quarter]
                analysis_period_keys = {
                    c for c in all_period_cols
                    if c[1:3] in q_months and c[3:] == filter_year
                }

        if analysis_period_keys:
            target_cols = sorted(
                [c for c in analysis_period_keys if c in all_period_cols],
                key=_period_sort_key,
            )
        else:
            target_cols = all_period_cols[-3:]

        if not target_cols:
            target_cols = all_period_cols[-3:]

        period_name_map: dict = {}
        params_path = _existing(period_id, "params")
        if params_path:
            try:
                saved = json.loads(params_path.read_text(encoding="utf-8"))
                period_name_map = saved.get("period_name_map", {})
            except Exception:
                pass
        period_labels_list = [period_name_map.get(c, c) for c in target_cols] or [period_label]

        analysis = await asyncio.get_event_loop().run_in_executor(
            None,
            analyze_cpi,
            df_for_analysis,
            target_cols,
            period_labels_list,
            period_label,
        )
        if analysis:
            try:
                await bot.send_message(chat_id, analysis, parse_mode="Markdown")
            except Exception:
                await bot.send_message(chat_id, analysis)
    except Exception as e:
        logger.warning(f"AI-анализ ИПЦ не удался: {e}")

async def check_cpi_new_periods(
    cookie: str,
    period_id: str,
) -> tuple[bool, set[str], dict[str, str]]:
    """
    Quick check: compare website periods against local CSV.
    Returns (ok, new_keys, period_map).
      new_keys   — column keys available on website but not in local CSV.
      period_map — {key: human_label} for every API period.
    """
    def _do():
        cpi_fetch  = _cpi_module("fetch")
        session    = cpi_fetch.make_session(cookie)
        params     = dict(cpi_fetch.KNOWN_PARAMS)
        params["period_id"] = period_id

        if not cpi_fetch.check_connection(session, params):
            raise RuntimeError("Cookie недействителен или сайт недоступен")

        _, date_list, name_list = cpi_fetch.get_period_names(session, params)
        api_keys   = {f"y{d}" for d in date_list}
        period_map = {f"y{d}": n for d, n in zip(date_list, name_list)}

        csv_path = _existing(period_id, "csv")
        if csv_path:
            cols       = pd.read_csv(csv_path, nrows=0, encoding="utf-8-sig").columns.tolist()
            local_keys = {c for c in cols if re.match(r"^y\d{6}$", c)}
        else:
            local_keys = set()

        return api_keys - local_keys, period_map

    try:
        new_keys, period_map = await asyncio.wait_for(
            asyncio.get_event_loop().run_in_executor(None, _do),
            timeout=120,
        )
        return True, new_keys, period_map
    except Exception as e:
        logger.error("CPI check error: %s", e)
        return False, set(), {}


async def run_cpi_check_update(
    chat_id: int,
    bot,
    cookie: str,
    period_id: str = "5",
) -> None:
    """
    Быстро проверяет наличие новых периодов на сайте.
    Если найдены новые данные — запускает полную загрузку.
    """
    period_label = PERIOD_LABELS.get(period_id, f"период {period_id}")

    ok, new_keys, period_map = await check_cpi_new_periods(cookie, period_id)

    if not ok:
        await bot.send_message(
            chat_id,
            "❌ Не удалось подключиться к сайту. Проверь Cookie.",
            parse_mode="Markdown",
        )
        return

    if not new_keys:
        await bot.send_message(
            chat_id,
            f"✅ *{period_label}* — данные уже актуальны.\nНовых периодов не найдено.",
            parse_mode="Markdown",
        )
        return

    new_names = sorted(period_map.get(k, k) for k in new_keys)
    await bot.send_message(
        chat_id,
        f"📊 Найдено *{len(new_keys)}* новых период(ов) — *{period_label}*:\n" +
        "\n".join(f"  • {n}" for n in new_names) +
        "\n\n⏳ Начинаю загрузку...",
        parse_mode="Markdown",
    )

    await run_cpi_pipeline(
        chat_id, bot, cookie,
        period_id=period_id,
        no_resume=True,
        analysis_period_keys=new_keys,
    )


# ── Filtered download helpers ──────────────────────────────────────────────────

def _period_cols(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if re.match(r"^y\d{6}$", c)]


def get_available_years(period_id: str) -> list[str]:
    """Returns sorted list of years (descending) available in the CSV."""
    csv_path = _existing(period_id, "csv")
    if csv_path is None:
        return []
    cols = pd.read_csv(csv_path, encoding="utf-8-sig", nrows=0).columns.tolist()
    return sorted({c[3:] for c in cols if re.match(r"^y\d{6}$", c)}, reverse=True)


def get_available_months(period_id: str, year: str) -> list[str]:
    """Returns sorted list of zero-padded month strings for a given year."""
    csv_path = _existing(period_id, "csv")
    if csv_path is None:
        return []
    cols = pd.read_csv(csv_path, encoding="utf-8-sig", nrows=0).columns.tolist()
    return sorted({c[1:3] for c in cols if re.match(r"^y\d{6}$", c) and c[3:] == year})


def build_filtered_excel(
    period_id: str,
    filter_type: str,
    year: str | None = None,
    month: str | None = None,
    quarter: int | None = None,
) -> tuple[Path, str]:
    """
    Builds a filtered Excel from the existing CSV.
    filter_type: "all" | "year" | "month" | "quarter"
    Returns (excel_path, human_label).
    """
    paths = _cpi_paths(period_id)

    # For "all" — send the existing full Excel if it's already built
    if filter_type == "all":
        existing_excel = _existing(period_id, "excel")
        if existing_excel:
            return existing_excel, "все периоды"

    csv_path = _existing(period_id, "csv")
    if csv_path is None:
        raise FileNotFoundError("CSV не найден")

    df = pd.read_csv(csv_path, encoding="utf-8-sig", low_memory=False)
    pcols = _period_cols(df)
    non_pcols = [c for c in df.columns if not re.match(r"^y\d{6}$", c)]

    if filter_type == "all":
        keep = pcols
        label = "все периоды"
    elif filter_type == "year":
        keep = [c for c in pcols if c[3:] == year]
        label = year
    elif filter_type == "month":
        keep = [c for c in pcols if c[1:3] == month and c[3:] == year]
        label = f"{MONTH_NAMES.get(int(month), month)} {year}"
    elif filter_type == "quarter":
        q_months = QUARTER_MONTHS[quarter]
        keep = [c for c in pcols if c[1:3] in q_months and c[3:] == year]
        label = f"Q{quarter} {year}"
    else:
        raise ValueError(f"Unknown filter_type: {filter_type}")

    if not keep:
        raise ValueError("Нет данных для выбранного периода")

    df_filtered = df[non_pcols + keep].dropna(subset=keep, how="all")

    out_path = _CPI_OUT / f"cpi_filtered_{period_id}.xlsx"
    cpi_excel = _cpi_module("build_excel")
    cpi_excel.build_excel(
        df_filtered,
        _existing(period_id, "params"),
        out_path,
    )
    return out_path, label
