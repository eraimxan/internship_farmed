import requests
import pandas as pd
import time

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": "https://taldau.stat.gov.kz/",
    "Cookie": "ASP.NET_SessionId=ssnyv1rzkllmvklnoieguopi; _ym_uid=1758390801129425250; _ym_d=1758390801"
})

BASE = "https://taldau.stat.gov.kz/ru/Api/GetIndexTreeData"
INDEX_ID = "701830"
PERIOD_ID = "7"
DIC_IDS = "68,776,59,3028,4043"
TERM_ID = "741885"

def fetch(p_terms, parent_id="", retries=3):
    params = {
        "p_measure_id": "1",
        "p_index_id": INDEX_ID,
        "p_period_id": PERIOD_ID,
        "p_terms": p_terms,
        "p_term_id": TERM_ID,
        "p_dicIds": DIC_IDS,
        "idx": "3",
        "p_parent_id": parent_id
    }
    for attempt in range(retries):
        try:
            r = session.get(BASE, params=params, timeout=60)
            time.sleep(0.5)
            if r.status_code == 200 and len(r.text.strip()) > 2:
                return r.json()
            return []
        except requests.exceptions.ReadTimeout:
            print(f"   Таймаут (попытка {attempt+1}/{retries}), жду 5 сек...")
            time.sleep(5)
        except Exception as e:
            print(f"   Ошибка: {e}")
            return []
    return []

def fetch_recursive(p_terms, parent_id="", depth=0, parent_text="", max_depth=5):
    if depth > max_depth:
        return []
    rows = fetch(p_terms, parent_id)
    all_rows = []
    for row in rows:
        row["depth"] = depth
        row["parent_text"] = parent_text
        all_rows.append(row)
        if row.get("leaf") == "false":
            children = fetch_recursive(
                p_terms,
                parent_id=row["id"],
                depth=depth + 1,
                parent_text=row["text"],
                max_depth=max_depth
            )
            all_rows.extend(children)
    return all_rows

akmola = [
    {"id": 247784, "name": "КОКШЕТАУ Г.А."},
    {"id": 247791, "name": "СТЕПНОГОРСК Г.А."},
    {"id": 247807, "name": "АККОЛЬСКИЙ РАЙОН"},
    {"id": 247855, "name": "АРШАЛЫНСКИЙ РАЙОН"},
    {"id": 247904, "name": "АСТРАХАНСКИЙ РАЙОН"},
    {"id": 247961, "name": "АТБАСАРСКИЙ РАЙОН"},
    {"id": 248026, "name": "БУЛАНДЫНСКИЙ РАЙОН"},
    {"id": 248076, "name": "ЕГИНДЫКОЛЬСКИЙ РАЙОН"},
    {"id": 248100, "name": "РАЙОН БИРЖАН САЛ"},
    {"id": 248165, "name": "ЕРЕЙМЕНТАУСКИЙ РАЙОН"},
    {"id": 248261, "name": "ЕСИЛЬСКИЙ РАЙОН"},
    {"id": 248322, "name": "ЖАКСЫНСКИЙ РАЙОН"},
    {"id": 248378, "name": "ЖАРКАИНСКИЙ РАЙОН"},
    {"id": 248433, "name": "ЗЕРЕНДИНСКИЙ РАЙОН"},
    {"id": 248558, "name": "КОРГАЛЖЫНСКИЙ РАЙОН"},
    {"id": 248592, "name": "САНДЫКТАУСКИЙ РАЙОН"},
    {"id": 248655, "name": "ЦЕЛИНОГРАДСКИЙ РАЙОН"},
    {"id": 248729, "name": "ШОРТАНДИНСКИЙ РАЙОН"},
    {"id": 248775, "name": "БУРАБАЙСКИЙ РАЙОН"},
    {"id": 77123369, "name": "КОСШЫ Г.А."},
]

all_data = []
for item in akmola:
    print(f"⏳ {item['name']}...")
    p_terms = f"{item['id']},741917,741907,741885,19202525"
    rows = fetch_recursive(p_terms, parent_id="")
    for row in rows:
        row["region"] = item["name"]
    all_data.extend(rows)
    print(f"  ✓ {len(rows)} строк")
    # Сохраняем промежуточно на случай обрыва
    pd.DataFrame(all_data).to_csv("akmola_full.csv", index=False, encoding="utf-8-sig")

result = pd.DataFrame(all_data)
result.to_csv("akmola_full.csv", index=False, encoding="utf-8-sig")
print(f"\n✅ Готово! Всего строк: {len(result)}")
