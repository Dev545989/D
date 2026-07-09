import sys
import json
import time
import requests
import random
from datetime import datetime

URL = "https://wd0ptz13zs-dsn.algolia.net/1/indexes/*/queries"

PARAMS = {
    "x-algolia-agent": "Algolia for JavaScript (4.24.0); Browser (lite)",
    "x-algolia-api-key": "cdd839b4fdac840289e88633779e8634",
    "x-algolia-application-id": "WD0PTZ13ZS",
}

CATEGORIES = {
    "rent_residential": {
        "index": "property-for-rent-residential.com",
        "filter": '("category_v2.slug_paths":"property-for-rent/residential")',
        "hits": []
    },
    "rent_commercial": {
        "index": "property-for-rent-commercial.com",
        "filter": '("category_v2.slug_paths":"property-for-rent/commercial")',
        "hits": []
    },
    "rent_rooms_rent_flatmates": {
        "index": "property-for-rent-rooms-for-rent-flatmates.com",
        "filter": '("category_v2.slug_paths":"property-for-rent/rooms-for-rent-flatmates")',
        "hits": []
    },
    "rent_holiday_homes": {
        "index": "property-for-rent-holiday-homes.com",
        "filter": '("categories.slug_paths":"property-for-rent/holiday-homes")',
        "hits": []
    },
    "rent_short_term_daily": {
        "index": "property-for-rent-short-term-daily.com",
        "filter": '("category_v2.slug_paths":"property-for-rent/short-term-daily")',
        "hits": []
    },
    "sale_residential": {
        "index": "property-for-sale-residential.com",
        "filter": '("category_v2.slug_paths":"property-for-sale/residential")',
        "hits": []
    },
    "sale_commercial": {
        "index": "property-for-sale-commercial.com",
        "filter": '("category_v2.slug_paths":"property-for-sale/commercial")',
        "hits": []
    },
    "jobs": {
        "index": "by_added_desc_jobs.com",
        "filter": '("category_v2.slug_paths":"jobs")',
        "hits": []
    },
    "jobs_wanted": {
        "index": "by_added_desc_jobs-wanted.com",
        "filter": '("category_v2.slug_paths":"jobs-wanted")',
        "hits": []
    }
}


def get_page_with_retry(category: dict, page: int, max_retries: int = 3) -> dict:
    payload = {
        "requests": [{
            "indexName": category["index"],
            "query": "",
            "params": f"page={page}&hitsPerPage=25&filters={category['filter']}",
        }]
    }

    for attempt in range(1, max_retries + 1):
        try:
            r = requests.post(URL, params=PARAMS, json=payload, timeout=30)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            print(f"  [Attempt {attempt}/{max_retries}] Page {page} failed: {e}")
            if attempt < max_retries:
                wait_time = attempt * 2
                print(f"  Waiting {wait_time}s before retry...")
                time.sleep(wait_time)

    return None


def run(category_name: str, start_page: int, end_page: int, output_jsonl: str) -> dict:
    if category_name not in CATEGORIES:
        print(f"Unknown category: {category_name}")
        return {"success": 0, "failed": 0, "failed_pages": [], "total_pages": 0}

    category = CATEGORIES[category_name]
    print(f"Scraping {category_name} | pages {start_page}-{end_page}")

    hits = []
    failed_pages = []
    total_pages = end_page - start_page + 1

    for page in range(start_page, end_page + 1):
        print(f"  Processing page {page}...")

        data = get_page_with_retry(category, page, max_retries=3)

        if data is None:
            print(f"  [FAILED] Page {page} failed after 3 attempts, skipping...")
            failed_pages.append(page)
            continue

        try:
            page_hits = data["results"][0]["hits"]
            print(f"  Page {page}: {len(page_hits)} listings")

            if not page_hits:
                print(f"  Page {page} has no results, stopping...")
                break

            hits.extend(page_hits)
            delay = random.uniform(0.5, 2.5)
            print(f"  Waiting {delay:.2f}s before next request...")
            time.sleep(delay)

        except Exception as e:
            print(f"  [ERROR] Page {page} data processing failed: {e}")
            failed_pages.append(page)

    with open(output_jsonl, "w", encoding="utf-8") as f:
        for hit in hits:
            f.write(json.dumps(hit, ensure_ascii=False) + "\n")

    if failed_pages:
        failed_file = output_jsonl.replace(".jsonl", "_failed.txt")
        with open(failed_file, "w", encoding="utf-8") as f:
            f.write(f"Category: {category_name}\n")
            f.write(f"Total pages in this job: {total_pages}\n")
            f.write(f"Failed pages: {len(failed_pages)}\n")
            f.write(f"Failed percentage: {(len(failed_pages)/total_pages)*100:.2f}%\n\n")
            for p in failed_pages:
                f.write(f"page={p}\n")

    print(f"Saved {len(hits)} listings to {output_jsonl} | {len(failed_pages)} failed pages")

    return {
        "success": len(hits),
        "failed": len(failed_pages),
        "failed_pages": failed_pages,
        "total_pages": total_pages
    }


if __name__ == "__main__":
    if len(sys.argv) == 4:
        category_name = sys.argv[1]
        start = int(sys.argv[2])
        end = int(sys.argv[3])
        output = f"{category_name}_{start}_{end}.jsonl"
        result = run(category_name, start, end, output)

        result_file = f"{category_name}_{start}_{end}_result.json"
        with open(result_file, "w", encoding="utf-8") as f:
            json.dump({
                **result,
                "category": category_name,
                "start_page": start,
                "end_page": end,
                "timestamp": datetime.now().isoformat()
            }, f, indent=2)
    else:
        print("Usage: python main.py <category_name> <start_page> <end_page>")
        sys.exit(1)