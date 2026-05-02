"""
Generic loader for Taldau (BNS АСПР) statistics API.
API base: https://taldau.stat.gov.kz/ru/Api/

Workflow per indicator:
  1) GetPeriodList(indexId)               -> available periods
  2) GetSegmentList(indexId, periodId)    -> available dimension segments (dics)
  3) GetIndexPeriods(...)                  -> available dates for that combination
  4) GetIndexTreeData(...)                 -> the actual values
  5) GetIndexAttributes(...)               -> indicator name/unit/decimal format
"""
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

API_BASE = "https://taldau.stat.gov.kz/ru/Api"
KATO_RK = 741880   # ID of РЕСПУБЛИКА КАЗАХСТАН
DEFAULT_TIMEOUT = 30

PERIOD_NAMES = {
    4: "Месяц",
    5: "Квартал",
    7: "Год",
    8: "Месяц с накоплением",
    9: "Квартал с накоплением",
}


def _get(url: str, params: dict = None) -> dict:
    r = requests.get(url, params=params, timeout=DEFAULT_TIMEOUT, verify=False)
    r.raise_for_status()
    if not r.text.strip():
        return None
    return r.json()


def get_period_list(index_id: int) -> list:
    """Available period types for an indicator: [{id, name}, ...]"""
    data = _get(f"{API_BASE}/GetPeriodList", {"indexId": index_id})
    return data or []


def get_segment_list(index_id: int, period_id: int) -> list:
    """Available dimension segments. Each item has dicId, termIds, idx, etc."""
    data = _get(f"{API_BASE}/GetSegmentList",
                {"indexId": index_id, "periodId": period_id})
    return data or []


def get_index_attributes(index_id: int, period_id: int) -> dict:
    """Indicator metadata: name, namePath, decFormat, measureName."""
    data = _get(f"{API_BASE}/GetIndexAttributes", {
        "indexId": index_id, "periodId": period_id,
        "measureID": 1, "measureKFC": 1,
    })
    return data or {}


def get_index_periods(index_id: int, period_id: int,
                      terms: str, term_id: int, dic_ids: str) -> dict:
    """Available dates for an index+segment combo.
    Returns dict with dateList, periodNameList, datesToDraw, datesIds."""
    data = _get(f"{API_BASE}/GetIndexPeriods", {
        "p_measure_id": 1,
        "p_index_id": index_id,
        "p_period_id": period_id,
        "p_terms": terms,
        "p_term_id": term_id,
        "p_dicIds": dic_ids,
    })
    return data or {}


def get_index_tree_data(index_id: int, period_id: int,
                        terms: str, term_id: int, dic_ids: str,
                        idx: int = 0, parent_id: str = "") -> list:
    """The actual data rows. Each row has fields y{dateKey} for each date."""
    data = _get(f"{API_BASE}/GetIndexTreeData", {
        "p_measure_id": 1,
        "p_index_id": index_id,
        "p_period_id": period_id,
        "p_terms": terms,
        "p_term_id": term_id,
        "p_dicIds": dic_ids,
        "idx": idx,
        "parent_id": parent_id,
    })
    return data or []


def pick_rk_segment(segments: list) -> dict | None:
    """Pick the segment containing whole-Kazakhstan KATO term (741880)."""
    for s in segments:
        term_ids = str(s.get("termIds", ""))
        if str(KATO_RK) in term_ids.replace(" ", "").split(","):
            return s
    return segments[0] if segments else None


def fetch_indicator(index_id: int, period_id: int) -> dict:
    """High-level fetch: returns {name, unit, dates, values, all_periods}.

    `dates` is list of (label, dd.mm.yyyy) tuples, `values` is parallel list of floats.
    """
    attrs = get_index_attributes(index_id, period_id)
    name = (attrs.get("name") or "").strip()
    unit = (attrs.get("measureName") or "").strip()

    segments = get_segment_list(index_id, period_id)
    seg = pick_rk_segment(segments)
    if not seg:
        return {"name": name, "unit": unit, "dates": [], "values": [], "rows": []}

    term_ids = str(seg.get("termIds", "")).replace(" ", "")
    dic_ids = str(seg.get("dicId", "")).replace(" ", "").replace("+", ",")
    idx = seg.get("idx", 0)

    periods = get_index_periods(index_id, period_id,
                                terms=term_ids,
                                term_id=KATO_RK,
                                dic_ids=dic_ids)
    date_list = periods.get("dateList", []) or []
    period_names = periods.get("periodNameList", []) or []
    dates_to_draw = periods.get("datesToDraw", []) or []

    rows = get_index_tree_data(index_id, period_id,
                               terms=term_ids,
                               term_id=KATO_RK,
                               dic_ids=dic_ids,
                               idx=idx)

    rk_row = next((r for r in rows if str(r.get("id")) == str(KATO_RK)), None)
    if rk_row is None and rows:
        rk_row = rows[0]

    dates_paired = []
    values = []
    for i, dkey in enumerate(date_list):
        label = period_names[i] if i < len(period_names) else dkey
        date_str = dates_to_draw[i] if i < len(dates_to_draw) else ""
        dates_paired.append((label.strip(), date_str.strip()))
        v = rk_row.get(f"y{dkey}") if rk_row else None
        try:
            values.append(float(str(v).replace(",", ".")) if v not in (None, "") else None)
        except Exception:
            values.append(None)

    return {
        "name": name,
        "unit": unit,
        "dates": dates_paired,
        "values": values,
        "rows": rows,
    }


def filter_by_year(dates: list, values: list, year: int) -> tuple[list, list]:
    """Keep only entries whose date string ends with the given year."""
    out_dates, out_values = [], []
    y_str = str(year)
    for (label, ds), v in zip(dates, values):
        if y_str in label or ds.endswith(y_str):
            out_dates.append((label, ds))
            out_values.append(v)
    return out_dates, out_values
