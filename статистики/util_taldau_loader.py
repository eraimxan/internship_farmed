import requests
import psycopg2

# --- Настройки базы данных ---
DB_CONFIG = {
    'host': 'localhost',
    'port': 5432,
    'database': 'ERI',
    'user': 'postgres',
    'password': '20041015R'
}

# --- URL для запроса ---
url = "https://taldau.stat.gov.kz/ru/Api/GetIndexData/701827?period=7&dics=68,776,90,4043"

# --- Получаем данные ---
response = requests.get(url)
response.raise_for_status()
data = response.json()

# --- Создаём таблицу, если нет ---
create_table_sql = """
CREATE TABLE IF NOT EXISTS investments_data (
    id SERIAL PRIMARY KEY,
    region VARCHAR(255),
    level1 VARCHAR(255),
    level2 VARCHAR(255),
    indicator VARCHAR(500),
    year INT,
    report_date DATE,
    value NUMERIC
);
"""

insert_sql = """
INSERT INTO investments_data (region, level1, level2, indicator, year, report_date, value)
VALUES (%s, %s, %s, %s, %s, %s, %s);
"""

# --- Подключение и запись ---
with psycopg2.connect(**DB_CONFIG) as conn:
    with conn.cursor() as cur:
        cur.execute(create_table_sql)
        conn.commit()

        for item in data:
            region, level1, level2, indicator = item["termNames"]

            for p in item["periods"]:
                year = int(p["name"].split()[0])  # например "2016 год" → 2016
                report_date = p["date"]
                value_str = p["value"]

                # --- Проверяем значение ---
                try:
                    value = float(value_str)
                except (ValueError, TypeError):
                    value = None  # если "x" или пусто, пишем NULL

                cur.execute(insert_sql, (region, level1, level2, indicator, year, report_date, value))

        conn.commit()

print("✅ Данные успешно загружены в таблицу investments_data!")
