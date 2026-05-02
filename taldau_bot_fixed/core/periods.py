"""
core/periods.py — Утилиты для работы с периодами (месяц/год).

Поддерживает два типа периодов:
  - Месячные (нарастающим итогом): y012025, y022025, ..., y122025
  - Годовые: y122019, y122020, ... (декабрь = полный год)

Формат ключа: y{MM}{YYYY}
"""

import re
from config.settings import MONTH_RU


def period_key(month: int, year: int) -> str:
    """y012024 — ключ колонки для январь 2024."""
    if not (1 <= month <= 12):
        raise ValueError(f"Месяц должен быть 1-12, получено: {month}")
    if not (2000 <= year <= 2100):
        raise ValueError(f"Год должен быть 2000-2100, получено: {year}")
    return f"y{month:02d}{year}"


def yearly_key(year: int) -> str:
    """y122024 — ключ годовой колонки (декабрь = весь год)."""
    return period_key(12, year)


def period_label(month: int, year: int) -> str:
    """
    Метка периода:
      month=1  → 'Янв 2025'
      month=6  → 'Янв-Июн 2025'
      month=12 → '2025 (год)'
    """
    if month == 12:
        return f"{year} (год)"
    if month == 1:
        return f"Янв {year}"
    return f"Янв-{MONTH_RU[month]} {year}"


def yearly_label(year: int) -> str:
    """'2024' — метка для годовых данных."""
    return str(year)


def parse_period_key(p: str) -> tuple[int, int] | None:
    """'y032024' → (3, 2024). None если не подходит формат."""
    m = re.match(r"^y(\d{2})(\d{4})$", p)
    if not m:
        return None
    month, year = int(m.group(1)), int(m.group(2))
    if not (1 <= month <= 12):
        return None
    return (month, year)


def fmt_period(p: str) -> str:
    """Красивое название периода из ключа колонки."""
    parsed = parse_period_key(p)
    return period_label(*parsed) if parsed else p


def period_sort_key(p: str) -> tuple[int, int]:
    """Ключ сортировки: (год, месяц)."""
    parsed = parse_period_key(p)
    if parsed:
        month, year = parsed
        return (year, month)
    return (0, 0)


def is_period_key(p: str) -> bool:
    """Возвращает True если ключ соответствует формату yMMYYYY."""
    return bool(re.match(r"^y\d{6}$", p))


def is_yearly_key(p: str) -> bool:
    """Возвращает True если это годовой период (декабрь = полный год)."""
    parsed = parse_period_key(p)
    return parsed is not None and parsed[0] == 12


def is_monthly_key(p: str) -> bool:
    """Возвращает True если это месячный период (не годовой)."""
    parsed = parse_period_key(p)
    return parsed is not None and parsed[0] != 12
