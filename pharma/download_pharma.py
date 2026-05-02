import requests
import json
import time
import os

BASE_URL = "https://pharma.eaeunion.org/spd2/find"
COLLECTION = "kbdallread.service-prop-26-drugRegistrationDetailsTypePublicList"
SORT = json.dumps({"resourceItemStatusDetails.updateDateTime": -1})
FIELDS = json.dumps({"id": 1, "drugCountryRegistrationDetails": 1, "drugDetails": 1})
LIMIT = 1000
OUTPUT_FILE = "pharma_data.json"


def fetch_batch(skip: int, retries: int = 3) -> tuple[list, int]:
    params = {
        "collection": COLLECTION,
        "limit": LIMIT,
        "skip": skip,
        "sort": SORT,
        "fields": FIELDS,
    }
    headers = {
        "Content-Type": "text/plain",
        "accept": "application/json, text/plain, */*",
    }

    for attempt in range(1, retries + 1):
        try:
            resp = requests.post(
                BASE_URL,
                headers=headers,
                data="{}",
                params=params,
                timeout=120,
            )
            resp.raise_for_status()
            data = resp.json()
            records = data.get("result", [])
            total = data.get("totalDocuments", 0)
            return records, total
        except requests.exceptions.Timeout:
            print(f"  Timeout on attempt {attempt}/{retries}, retrying...", flush=True)
            time.sleep(5 * attempt)
        except requests.exceptions.RequestException as e:
            print(f"  Request error on attempt {attempt}/{retries}: {e}", flush=True)
            time.sleep(5 * attempt)

    return [], 0


def main():
    all_records = []
    skip = 0
    batch_num = 0
    total_documents = None

    print("Downloading EAEU Unified Register of Registered Medicines...")
    print(f"Collection: {COLLECTION}")
    print(f"Batch size: {LIMIT}")
    print("-" * 60)

    while True:
        batch_num += 1
        print(f"Batch {batch_num}: records {skip + 1}–{skip + LIMIT}...", end=" ", flush=True)

        batch, total = fetch_batch(skip=skip)

        if total_documents is None and total:
            total_documents = total
            print(f"\n  Total documents: {total_documents}")
            print(f"  Batches needed: {(total_documents + LIMIT - 1) // LIMIT}")

        if not batch:
            print("Empty response, stopping.")
            break

        all_records.extend(batch)
        print(f"got {len(batch)}. Total: {len(all_records)}/{total_documents}", flush=True)

        if len(batch) < LIMIT:
            print("Last batch received.")
            break

        skip += LIMIT
        time.sleep(0.3)

    print("-" * 60)
    print(f"Download complete. Total records: {len(all_records)}")

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(all_records, f, ensure_ascii=False, indent=2)

    size_mb = os.path.getsize(OUTPUT_FILE) / (1024 * 1024)
    print(f"Saved to {OUTPUT_FILE} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
