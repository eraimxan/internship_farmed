"""
ai_analyzer.py — Анализ новых данных через Google Gemini API.
Поддерживает два типа данных: ИОК и ИПЦ.

Готовит подробный текстовый дайджест по данным (общая инфляция/инвестиции,
основные группы, топ-движители, региональный разброс) и просит модель
сформировать пост в стиле TENGENOMIKA.
"""

import logging
import os
import re

import pandas as pd
import requests

logger = logging.getLogger(__name__)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent"


# ─────────────────────────────────────────────────────────────────────
# ВСПОМОГАТЕЛЬНОЕ
# ─────────────────────────────────────────────────────────────────────

def _period_sort_key(col: str) -> tuple[int, int]:
    """Chronological sort for yMMYYYY columns (e.g. y012025 → (2025, 1))."""
    m = re.match(r"^y(\d{2})(\d{4})$", col)
    if m:
        return (int(m.group(2)), int(m.group(1)))
    return (0, 0)


def _sort_period_pairs(cols: list, labels: list) -> tuple[list, list]:
    """Return cols and labels both sorted chronologically."""
    pairs = sorted(zip(cols, labels), key=lambda x: _period_sort_key(x[0]))
    return [p[0] for p in pairs], [p[1] for p in pairs]


def _fmt_pct(val: float, sign: bool = True) -> str:
    """11.83 → '+11.8%' (CPI delta from index)."""
    if pd.isna(val):
        return "н/д"
    s = f"{val:+.1f}" if sign else f"{val:.1f}"
    return f"{s}%"


def _fmt_money(val: float) -> str:
    """Pretty number with thousand separators."""
    if pd.isna(val):
        return "н/д"
    return f"{val:,.0f}".replace(",", " ")


def _period_unit(period_type: str) -> str:
    """Map Russian period type to lowercase unit ('квартал', 'месяц')."""
    pt = (period_type or "").lower()
    if "квартал" in pt:
        return "квартал"
    if "накоплен" in pt:
        return "месяц с накоплением"
    if "месяц" in pt:
        return "месяц"
    if "год" in pt:
        return "год"
    return pt or "период"


# ─────────────────────────────────────────────────────────────────────
# ШАБЛОНЫ
# ─────────────────────────────────────────────────────────────────────

CPI_TEMPLATE = """Ты — экономический аналитик в стиле TENGENOMIKA. Пишешь пост для Telegram-канала об инфляции в Казахстане.

Тип периода: {period_type}. Период: {last_period}.
В данных ниже все числа уже приведены к виду «прирост в %» (например, +11.7% — это инфляция год к году).
Анализируй ИМЕННО эти данные — не выдумывай регионы, цифры и категории.

============ ДАННЫЕ ============
{data_summary}
================================

ТРЕБОВАНИЯ К ВЫВОДУ:
1. Жирный текст ТОЛЬКО через одиночные звёздочки: *текст* — НЕ **текст**, никаких ## или _курсива_.
2. Используй ТОЛЬКО цифры и названия из данных выше.
3. Пиши на русском, в деловом аналитическом тоне (как TENGENOMIKA).
4. Объём: 250–400 слов. Не короче.
5. Формат — строго по разделам ниже, с теми же эмодзи и пунктуацией.

ВЫВЕДИ В ЭТОМ ФОРМАТЕ (подставь реальные числа из данных, скобки убери):

*Инфляция в Казахстане — {last_period}*
(1–2 предложения: уровень годовой инфляции, сравнение с предыдущим {period_unit} — ускорилась/замедлилась/сохранилась; динамика к предыдущему {period_unit_short} если есть)

📌 *Основные цифры по годовой инфляции:*
▫️Продукты: +X% (+Y% в пред. {period_unit_short})
▫️Непродтовары: +X% (+Y% в пред. {period_unit_short})
▫️Услуги: +X% (+Y% в пред. {period_unit_short}) — (короткий комментарий, если есть резкое изменение)

📍 *Что подорожало сильнее всего за год:*
▪️(Категория): +X%
▪️(Категория): +X%
▪️(Категория): +X%
▪️(Категория): +X%
▪️(Категория): +X%
▪️(Категория): +X%
▪️(Категория): +X%
(используй 6–8 позиций из блока «Топ движителей»)

❗️ *Региональный разброс:*
▫️Самая высокая инфляция: (Регион) (+X%), (Регион) (+Y%), (Регион) (+Z%)
▫️Самая низкая: (Регион) (+X%), (Регион) (+Y%)

💬 (3–5 предложений: что означает текущая динамика, основные драйверы (тарифы ЖКХ / продовольствие / услуги), региональная асимметрия, краткосрочный прогноз. Тон — аналитический, без воды.)
"""


IOK_TEMPLATE = """Ты — экономический аналитик в стиле TENGENOMIKA. Пишешь пост для Telegram-канала об инвестициях в основной капитал (ИОК) в Казахстане.

Период: {last_period}.
Все суммы в данных ниже приведены в тыс. тенге.
Все приросты в % уже посчитаны.

============ ДАННЫЕ ============
{data_summary}
================================

ТРЕБОВАНИЯ К ВЫВОДУ:
1. Жирный текст ТОЛЬКО через одиночные звёздочки: *текст* — НЕ **текст**, никаких ## или _курсива_.
2. Используй ТОЛЬКО цифры, регионы и категории из данных выше.
3. Пиши на русском, в деловом аналитическом тоне (как TENGENOMIKA).
4. Объём: 250–400 слов. Не короче.
5. Формат — строго по разделам ниже, с теми же эмодзи.

ВЫВЕДИ В ЭТОМ ФОРМАТЕ (подставь реальные числа, скобки убери):

*Инвестиции в основной капитал — {last_period}*
(1–2 предложения: суммарный объём по Казахстану, динамика к предыдущему периоду — ускорился/замедлился/стабилен, концентрация в топ-3 регионах)

📌 *Основные цифры по структуре инвестиций:*
▫️(Категория 1): (сумма) тыс. тенге (+X% к пред. периоду)
▫️(Категория 2): (сумма) тыс. тенге (+X% к пред. периоду)
▫️(Категория 3): (сумма) тыс. тенге (+X% к пред. периоду)
▫️(Категория 4): (сумма) тыс. тенге (+X%)
(используй блок «Основные категории инвестиций» — 4–6 позиций)

📍 *Лидеры роста по регионам:*
▪️(Область): +X% к предыдущему периоду
▪️(Область): +X%
▪️(Область): +X%
▪️(Область): +X%
▪️(Область): +X%

❗️ *Региональный разброс:*
▫️Топ по объёму: (Область) — (сумма) тыс. тенге, (Область) — (сумма), (Область) — (сумма)
▫️Минимальный объём: (Область) — (сумма) тыс. тенге, (Область) — (сумма)
▫️Падение/слабая динамика: (Область) (X%), (Область) (X%)

💬 (3–5 предложений: ключевые факторы динамики, что объясняет региональную концентрацию, какие отрасли тянут / тормозят, что это означает для структуры экономики, краткосрочный прогноз. Тон — аналитический.)
"""


# ─────────────────────────────────────────────────────────────────────
# ПОДГОТОВКА ДАННЫХ ИПЦ
# ─────────────────────────────────────────────────────────────────────

# Канонические тексты в данных Talдау
_OVERALL_CPI_TEXTS = {"товары и услуги"}
_GROUP_CPI_TEXTS = {
    "продовольственные товары":   "Продукты",
    "непродовольственные товары": "Непродтовары",
    "платные услуги":             "Услуги",
}
# Аггрегаты depth=0, которые не нужны в «топ движителях» (это широкие группы)
_AGG_TEXTS = {
    "товары и услуги",
    "товары",
    "продовольственные товары",
    "непродовольственные товары",
    "платные услуги",
    "продукты питания и товары для младенцев",
    "продукты питания без фруктов и овощей",
    "все товары и услуги без фруктов и овощей",
    "товары для детей школьного возраста",
}
_REPUBLIC = "РЕСПУБЛИКА КАЗАХСТАН"


def _coerce_periods(df: pd.DataFrame, period_cols: list) -> pd.DataFrame:
    out = df.copy()
    for c in period_cols:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce")
    return out


def _pretty_oblast(name: str) -> str:
    """'АЛМАТИНСКАЯ ОБЛАСТЬ' → 'Алматинская область'; 'Г.АЛМАТЫ' → 'г. Алматы'."""
    if not isinstance(name, str):
        return str(name)
    n = name.strip()
    if n.upper().startswith("Г."):
        rest = n[2:].strip().capitalize()
        return f"г. {rest}"
    return n.capitalize()


def _prepare_cpi_summary(
    df: pd.DataFrame,
    period_cols: list,
    period_labels: list,
) -> str:
    """
    Богатый дайджест ИПЦ:
      • Общий ИПЦ (Республика, г/г и к пред. периоду)
      • Основные группы (Продукты/Непродтовары/Услуги) — текущий и предыдущий периоды
      • Топ-движители (конкретные позиции, г/г)
      • Региональная картина (общий ИПЦ по областям, г/г)
    """
    if df.empty or not period_cols:
        return "Данные недоступны."

    has_comp = "comparison_type" in df.columns
    has_depth = "depth" in df.columns
    has_oblast = "oblast" in df.columns
    has_text = "text" in df.columns
    if not (has_comp and has_depth and has_oblast and has_text):
        return "Данные недоступны (отсутствуют ключевые столбцы)."

    df = _coerce_periods(df, period_cols)
    # depth → str безопасно
    df["_depth"] = df["depth"].astype(str)
    df["_text_l"] = df["text"].astype(str).str.strip().str.lower()

    last_col, last_label = period_cols[-1], period_labels[-1]
    prev_col, prev_label = (
        (period_cols[-2], period_labels[-2]) if len(period_cols) >= 2 else (None, None)
    )

    yoy = df[df["comparison_type"] == "yoy"]
    pp  = df[df["comparison_type"] == "prev_period"]

    rep_yoy = yoy[yoy["oblast"] == _REPUBLIC]
    rep_pp  = pp[pp["oblast"]  == _REPUBLIC]

    lines: list[str] = []
    lines.append(f"ПЕРИОД: {last_label}")
    if prev_label:
        lines.append(f"ПРЕДЫДУЩИЙ ПЕРИОД В ДАННЫХ: {prev_label}")
    lines.append("")

    # ── Общая инфляция ───────────────────────────────────────────────
    def _val(sub: pd.DataFrame, text_l: str, col: str) -> float:
        if not col or col not in sub.columns:
            return float("nan")
        match = sub[(sub["_text_l"] == text_l) & (sub["_depth"] == "0")]
        if match.empty:
            return float("nan")
        v = match.iloc[0][col]
        return (v - 100) if pd.notna(v) else float("nan")

    lines.append("ОБЩАЯ ИНФЛЯЦИЯ (Республика Казахстан):")
    overall_yoy_last = _val(rep_yoy, "товары и услуги", last_col)
    overall_pp_last  = _val(rep_pp,  "товары и услуги", last_col)
    lines.append(f"  Годовая (г/г) в {last_label}: {_fmt_pct(overall_yoy_last)}")
    lines.append(f"  К предыдущему периоду в {last_label}: {_fmt_pct(overall_pp_last)}")
    if prev_col:
        overall_yoy_prev = _val(rep_yoy, "товары и услуги", prev_col)
        overall_pp_prev  = _val(rep_pp,  "товары и услуги", prev_col)
        lines.append(f"  Годовая (г/г) в {prev_label}: {_fmt_pct(overall_yoy_prev)}")
        lines.append(f"  К предыдущему периоду в {prev_label}: {_fmt_pct(overall_pp_prev)}")
        if pd.notna(overall_yoy_last) and pd.notna(overall_yoy_prev):
            delta = overall_yoy_last - overall_yoy_prev
            direction = "ускорилась" if delta > 0.05 else ("замедлилась" if delta < -0.05 else "сохранилась")
            lines.append(f"  Динамика годовой инфляции: {direction} ({delta:+.1f} п.п.)")
    lines.append("")

    # ── Основные группы ──────────────────────────────────────────────
    lines.append("ОСНОВНЫЕ ГРУППЫ (г/г, Республика Казахстан):")
    for txt_l, label in _GROUP_CPI_TEXTS.items():
        cur = _val(rep_yoy, txt_l, last_col)
        line = f"  {label}: {_fmt_pct(cur)} в {last_label}"
        if prev_col:
            prv = _val(rep_yoy, txt_l, prev_col)
            line += f" (в {prev_label}: {_fmt_pct(prv)})"
        lines.append(line)
    lines.append("")

    # ── Топ движителей ────────────────────────────────────────────────
    movers = rep_yoy[
        (rep_yoy["_depth"].isin(["1", "2", "3", "4", "5", "6", "7"]))
        & (~rep_yoy["_text_l"].isin(_AGG_TEXTS))
        & (rep_yoy[last_col].notna())
    ].copy()
    movers["_pct"] = movers[last_col] - 100
    movers = movers.sort_values("_pct", ascending=False)
    # Дедуп по тексту: иногда одна и та же категория встречается на разных глубинах
    movers = movers.drop_duplicates(subset=["_text_l"], keep="first")

    lines.append(f"ТОП ДВИЖИТЕЛЕЙ (рост цен г/г в {last_label}, конкретные позиции):")
    for _, r in movers.head(15).iterrows():
        lines.append(f"  {r['text']}: {_fmt_pct(r['_pct'])}")
    lines.append("")

    lines.append(f"АНТИЛИДЕРЫ (минимальный рост / снижение г/г в {last_label}):")
    bottom = movers.sort_values("_pct").head(5)
    for _, r in bottom.iterrows():
        lines.append(f"  {r['text']}: {_fmt_pct(r['_pct'])}")
    lines.append("")

    # ── Региональная картина ────────────────────────────────────────
    reg_yoy = yoy[
        (yoy["oblast"] != _REPUBLIC)
        & (yoy["_depth"] == "0")
        & (yoy["_text_l"] == "товары и услуги")
        & (yoy[last_col].notna())
    ].copy()
    if not reg_yoy.empty:
        reg_yoy["_pct"] = reg_yoy[last_col] - 100
        reg_yoy = (
            reg_yoy.groupby("oblast", as_index=False)["_pct"].mean()
            .sort_values("_pct", ascending=False)
        )
        lines.append(f"РЕГИОНАЛЬНАЯ КАРТИНА (общий ИПЦ г/г в {last_label}):")
        lines.append("  Высокая инфляция:")
        for _, r in reg_yoy.head(5).iterrows():
            lines.append(f"    {_pretty_oblast(r['oblast'])}: {_fmt_pct(r['_pct'])}")
        lines.append("  Низкая инфляция:")
        for _, r in reg_yoy.tail(5).iloc[::-1].iterrows():
            lines.append(f"    {_pretty_oblast(r['oblast'])}: {_fmt_pct(r['_pct'])}")

    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────
# ПОДГОТОВКА ДАННЫХ ИОК
# ─────────────────────────────────────────────────────────────────────

def _prepare_iok_summary(
    df: pd.DataFrame,
    period_cols: list,
    period_labels: list,
) -> str:
    """
    Богатый дайджест ИОК (как у CPI):
      • Суммарный объём + концентрация (доля топ-3)
      • Топ-10 областей по объёму с приростом
      • Лидеры/антилидеры роста
      • Региональный разброс (max/min)
      • Основные категории инвестиций (depth=1) с динамикой
      • Топ-движители (depth>=2: конкретные сектора с наибольшим ростом)
    """
    if df.empty or not period_cols:
        return "Данные недоступны."

    if "depth" not in df.columns or "oblast" not in df.columns:
        return "Данные недоступны (отсутствуют ключевые столбцы)."

    df = _coerce_periods(df, period_cols)
    df["_depth"] = df["depth"].astype(str)

    last_col, last_label = period_cols[-1], period_labels[-1]
    prev_col, prev_label = (
        (period_cols[-2], period_labels[-2]) if len(period_cols) >= 2 else (None, None)
    )

    # depth=0 «Всего» — итог по району, нужно суммировать на уровень области.
    base = df[df["_depth"] == "0"].copy()
    if base.empty:
        return "Нет агрегатов уровня depth=0."

    # тыс. тенге
    base["_last"] = base[last_col] / 1000 if last_col in base.columns else float("nan")
    if prev_col and prev_col in base.columns:
        base["_prev"] = base[prev_col] / 1000
    else:
        base["_prev"] = float("nan")

    by_oblast = (
        base.groupby("oblast", as_index=False)
        .agg(_last=("_last", "sum"), _prev=("_prev", "sum"))
    )
    by_oblast = by_oblast[by_oblast["_last"].notna() & (by_oblast["_last"] > 0)].copy()
    if by_oblast.empty:
        return "Нет числовых данных по областям."

    by_oblast["_growth"] = by_oblast.apply(
        lambda r: ((r["_last"] - r["_prev"]) / r["_prev"] * 100)
        if (pd.notna(r["_prev"]) and r["_prev"] > 0)
        else float("nan"),
        axis=1,
    )
    by_oblast = by_oblast.sort_values("_last", ascending=False)

    total_last = by_oblast["_last"].sum()
    total_prev = by_oblast["_prev"].sum() if by_oblast["_prev"].notna().any() else float("nan")

    lines: list[str] = []
    lines.append(f"ПЕРИОД: {last_label}")
    if prev_label:
        lines.append(f"ПРЕДЫДУЩИЙ ПЕРИОД В ДАННЫХ: {prev_label}")
    lines.append("")

    lines.append("СУММАРНЫЙ ОБЪЁМ ПО КАЗАХСТАНУ (тыс. тенге):")
    lines.append(f"  {last_label}: {_fmt_money(total_last)}")
    if pd.notna(total_prev) and total_prev > 0:
        growth = (total_last - total_prev) / total_prev * 100
        direction = "ускорился" if growth > 0.5 else ("замедлился" if growth < -0.5 else "стабилен")
        lines.append(f"  {prev_label}: {_fmt_money(total_prev)}")
        lines.append(f"  Прирост: {_fmt_pct(growth)} ({direction})")
    # Концентрация
    top3_share = by_oblast.head(3)["_last"].sum() / total_last * 100 if total_last > 0 else float("nan")
    top5_share = by_oblast.head(5)["_last"].sum() / total_last * 100 if total_last > 0 else float("nan")
    lines.append(f"  Доля топ-3 областей: {_fmt_pct(top3_share, sign=False)}")
    lines.append(f"  Доля топ-5 областей: {_fmt_pct(top5_share, sign=False)}")
    lines.append("")

    lines.append("ТОП-10 ОБЛАСТЕЙ ПО ОБЪЁМУ (тыс. тенге):")
    for _, r in by_oblast.head(10).iterrows():
        gtxt = f", прирост {_fmt_pct(r['_growth'])}" if pd.notna(r["_growth"]) else ""
        lines.append(f"  {_pretty_oblast(r['oblast'])}: {_fmt_money(r['_last'])}{gtxt}")
    lines.append("")

    growth_only = by_oblast.dropna(subset=["_growth"]).copy()
    if not growth_only.empty:
        lines.append("ЛИДЕРЫ РОСТА ПО РЕГИОНАМ (к предыдущему периоду):")
        for _, r in growth_only.sort_values("_growth", ascending=False).head(5).iterrows():
            lines.append(f"  {_pretty_oblast(r['oblast'])}: {_fmt_pct(r['_growth'])}")
        lines.append("")
        lines.append("АНТИЛИДЕРЫ (падение / минимальный рост):")
        for _, r in growth_only.sort_values("_growth").head(5).iterrows():
            lines.append(f"  {_pretty_oblast(r['oblast'])}: {_fmt_pct(r['_growth'])}")
        lines.append("")

    lines.append("РЕГИОНАЛЬНЫЙ РАЗБРОС:")
    top3 = by_oblast.head(3)
    bot3 = by_oblast.tail(3).iloc[::-1]
    lines.append("  Наибольший объём:")
    for _, r in top3.iterrows():
        lines.append(f"    {_pretty_oblast(r['oblast'])}: {_fmt_money(r['_last'])} тыс. тенге")
    lines.append("  Наименьший объём:")
    for _, r in bot3.iterrows():
        lines.append(f"    {_pretty_oblast(r['oblast'])}: {_fmt_money(r['_last'])} тыс. тенге")
    lines.append("")

    # ── Категории инвестиций (depth=1) с динамикой ───────────────────
    if "text" in df.columns:
        cats = df[df["_depth"] == "1"].copy()
        if not cats.empty and last_col in cats.columns:
            cats["_last"] = cats[last_col] / 1000
            if prev_col and prev_col in cats.columns:
                cats["_prev"] = cats[prev_col] / 1000
            else:
                cats["_prev"] = float("nan")
            cat_tot = (
                cats.groupby("text", as_index=False)
                .agg(_last=("_last", "sum"), _prev=("_prev", "sum"))
            )
            cat_tot = cat_tot[cat_tot["_last"] > 0].copy()
            cat_tot["_growth"] = cat_tot.apply(
                lambda r: ((r["_last"] - r["_prev"]) / r["_prev"] * 100)
                if (pd.notna(r["_prev"]) and r["_prev"] > 0)
                else float("nan"),
                axis=1,
            )
            cat_tot = cat_tot.sort_values("_last", ascending=False).head(8)
            if not cat_tot.empty:
                lines.append("ОСНОВНЫЕ КАТЕГОРИИ ИНВЕСТИЦИЙ (Республика, тыс. тенге):")
                for _, r in cat_tot.iterrows():
                    gtxt = f" (прирост {_fmt_pct(r['_growth'])})" if pd.notna(r["_growth"]) else ""
                    lines.append(f"  {r['text']}: {_fmt_money(r['_last'])}{gtxt}")
                lines.append("")

        # ── Топ-движители: конкретные сектора (depth >= 2) ───────────
        deep = df[df["_depth"].isin(["2", "3", "4", "5", "6", "7"])].copy()
        if not deep.empty and last_col in deep.columns:
            deep["_last"] = deep[last_col] / 1000
            if prev_col and prev_col in deep.columns:
                deep["_prev"] = deep[prev_col] / 1000
            else:
                deep["_prev"] = float("nan")
            sec = (
                deep.groupby("text", as_index=False)
                .agg(_last=("_last", "sum"), _prev=("_prev", "sum"))
            )
            sec = sec[sec["_last"] > 0].copy()
            sec["_growth"] = sec.apply(
                lambda r: ((r["_last"] - r["_prev"]) / r["_prev"] * 100)
                if (pd.notna(r["_prev"]) and r["_prev"] > 0)
                else float("nan"),
                axis=1,
            )
            growers = sec.dropna(subset=["_growth"])
            # отсечь шум: значимый объём + прежняя база не должна быть микроскопической
            # (иначе прирост от нуля даёт абсурдные тысячи %)
            growers = growers[
                (growers["_last"] > total_last * 0.002)
                & (growers["_prev"] > total_prev * 0.0005 if pd.notna(total_prev) else True)
                & (growers["_growth"] < 500)
                & (growers["_growth"] > -95)
            ]
            if not growers.empty:
                top_g = growers.sort_values("_growth", ascending=False).head(10)
                lines.append("ТОП-ДВИЖИТЕЛИ (сектора с наибольшим ростом инвестиций):")
                for _, r in top_g.iterrows():
                    lines.append(f"  {r['text']}: {_fmt_pct(r['_growth'])} (объём {_fmt_money(r['_last'])})")
                lines.append("")
                bot_g = growers.sort_values("_growth").head(5)
                lines.append("СОКРАЩЕНИЕ ИНВЕСТИЦИЙ (сектора с наибольшим падением):")
                for _, r in bot_g.iterrows():
                    lines.append(f"  {r['text']}: {_fmt_pct(r['_growth'])} (объём {_fmt_money(r['_last'])})")

    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────
# ВЫЗОВ GEMINI API
# ─────────────────────────────────────────────────────────────────────

def _call_gemini(prompt: str) -> str | None:
    if not GEMINI_API_KEY:
        logger.warning("GEMINI_API_KEY не задан — анализ пропущен")
        return None

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.4,
            "maxOutputTokens": 8192,
            "thinkingConfig": {"thinkingBudget": 0},
        },
    }
    try:
        resp = requests.post(
            GEMINI_URL,
            params={"key": GEMINI_API_KEY},
            json=payload,
            timeout=120,
        )
        resp.raise_for_status()
        data = resp.json()
        cand = data["candidates"][0]
        finish = cand.get("finishReason", "")
        parts = cand.get("content", {}).get("parts", []) or []
        text = "".join(p.get("text", "") for p in parts).strip()
        if finish and finish not in ("STOP", "MODEL_LENGTH"):
            logger.warning("Gemini finishReason=%s, len=%d", finish, len(text))
        if not text:
            logger.error("Gemini вернул пустой текст (finishReason=%s)", finish)
            return None
        return text
    except requests.exceptions.Timeout:
        logger.error("Gemini API: таймаут")
    except requests.exceptions.HTTPError as e:
        logger.error("Gemini API HTTP ошибка: %s", e)
    except (KeyError, IndexError) as e:
        logger.error("Gemini API: неожиданный ответ: %s", e)
    except Exception as e:
        logger.error("Gemini API ошибка: %s", e)
    return None


# ─────────────────────────────────────────────────────────────────────
# ПУБЛИЧНЫЕ ФУНКЦИИ
# ─────────────────────────────────────────────────────────────────────

def analyze_iok(
    master_df: pd.DataFrame,
    new_period_cols: list,
    period_labels: list,
) -> str | None:
    """Анализ данных ИОК (инвестиции в основной капитал)."""
    if master_df.empty or not new_period_cols:
        return None

    cols_sorted, labels_sorted = _sort_period_pairs(new_period_cols, period_labels)
    summary = _prepare_iok_summary(master_df, cols_sorted, labels_sorted)
    last_label = labels_sorted[-1] if labels_sorted else "период"
    prompt = IOK_TEMPLATE.format(last_period=last_label, data_summary=summary)
    return _call_gemini(prompt)


def analyze_cpi(
    df: pd.DataFrame,
    new_period_cols: list,
    period_labels: list,
    period_type: str = "Квартал",
) -> str | None:
    """Анализ данных ИПЦ (индекс потребительских цен)."""
    if df.empty or not new_period_cols:
        return None

    cols_sorted, labels_sorted = _sort_period_pairs(new_period_cols, period_labels)
    summary = _prepare_cpi_summary(df, cols_sorted, labels_sorted)
    last_label = labels_sorted[-1] if labels_sorted else "период"
    unit = _period_unit(period_type)
    # «месяц с накоплением» в скобках мелким — для коротких упоминаний используем «месяц»
    unit_short = "месяц" if "месяц" in unit else unit
    prompt = CPI_TEMPLATE.format(
        period_type=period_type,
        last_period=last_label,
        period_unit=unit,
        period_unit_short=unit_short,
        data_summary=summary,
    )
    return _call_gemini(prompt)
