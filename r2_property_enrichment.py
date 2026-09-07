"""
R2 Property phone + description enrichment.

Unlike classifieds/community/motors, a property listing's phone number is
NOT obtained by visiting the listing's own page -- it's obtained by visiting
the *agent* or *agency* profile page that posted the listing (a single
agent/agency can post hundreds of listings, so the same profile page would
otherwise get visited over and over).

Pipeline:
  1) prepare : read today's property Excel files from R2, read the day's
     cached profiles-data.xlsx (phone per agent/agency slug), and build:
       - "description" jobs: one per listing missing description_full
       - "profile" jobs: one per UNIQUE (profile_type, slug) missing a phone
         and not already cached (deduplicated across ALL listings this run)
  2) scrape : each job item is either kind="description" (visit the listing
     page) or kind="profile" (visit the agent/agency profile page).
  3) combine-profiles : merge new profile phones into profiles-data.xlsx on R2.
  4) combine : merge new profile phones into the cached profiles-data.xlsx,
     then re-read EVERY property file for the day and, for every row, fill
     in description_full (from its own scrape result) and contact_phone_number
     (looked up from the merged profiles cache by the row's agent/agency slug)
     -- even for rows whose own file had no scrape job at all.
"""

from __future__ import annotations

import argparse
import ast
import glob
import io
import json
import os
import random
import re
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import boto3
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

R2_BUCKET = os.environ["CF_R2_BUCKET_NAME"]
R2_ENDPOINT = os.environ["CF_R2_ENDPOINT_URL"]
R2_ACCESS_KEY = os.environ["CF_R2_ACCESS_KEY_ID"]
R2_SECRET_KEY = os.environ["CF_R2_SECRET_ACCESS_KEY"]

# NOTE: the property scraper (rent_property.yml -> final_merge) saves under
# f"DUAE/{date_prefix}/{category_path}/..." where date_prefix is built from
# datetime.now() with NO explicit timezone -- i.e. the runner's local/UTC date
# on the day the scraper actually ran, not "yesterday". So when THIS script
# runs the next day, "yesterday" (UTC) should line up with that save date.
# Verify against a real R2 run and pass --date explicitly if it doesn't match.
DUAE_PREFIX = os.getenv("MOTORS_PREFIX", "DUAE")
PROPERTY_ROOT = "property"

PHONE_COLUMN = "contact_phone_number"
DESCRIPTION_COLUMN = "description_full"
ABS_URL_COLUMN = "absolute_url"
AGENT_PROFILE_COLUMN = "agent_profile"
AGENT_COLUMN = "agent"

PROFILE_BASE_URLS = {
    "agency": "https://uae.dubizzle.com/property-agencies/{slug}/",
    "agent": "https://uae.dubizzle.com/property-agents/{slug}/",
}


def r2_client():
    return boto3.client(
        "s3",
        endpoint_url=R2_ENDPOINT,
        aws_access_key_id=R2_ACCESS_KEY,
        aws_secret_access_key=R2_SECRET_KEY,
        region_name="auto",
    )


def yesterday_prefix(date_str: str | None = None) -> tuple[str, str]:
    if date_str:
        target = datetime.strptime(date_str, "%Y-%m-%d").date()
    else:
        target = (datetime.now(timezone.utc) - timedelta(days=1)).date()
    prefix = f"{DUAE_PREFIX}/year={target.year}/month={target.month:02d}/day={target.day:02d}/"
    return target.isoformat(), prefix


def list_keys(client, prefix: str) -> list[str]:
    keys = []
    paginator = client.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=R2_BUCKET, Prefix=prefix):
        for obj in page.get("Contents", []):
            key = obj["Key"]
            if key.lower().endswith(".xlsx") and "/profiles-data/" not in key:
                keys.append(key)
    return sorted(keys)


def download_bytes(client, key: str) -> bytes:
    return client.get_object(Bucket=R2_BUCKET, Key=key)["Body"].read()


def upload_bytes(client, key: str, data: bytes, content_type: str):
    client.put_object(Bucket=R2_BUCKET, Key=key, Body=data, ContentType=content_type)


def is_empty(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip().lower() in {"", "nan", "none", "null"}


def clean_excel_value(value):
    if isinstance(value, str):
        return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", value)
    return value


def excel_sheets(data: bytes) -> dict[str, pd.DataFrame]:
    # dtype=str on the phone column is critical: otherwise pandas infers an
    # int64 column for all-digit phone numbers and silently drops the
    # leading zero (e.g. "0501234567" -> 501234567) every time the file is
    # read back.
    return pd.read_excel(io.BytesIO(data), sheet_name=None, dtype={PHONE_COLUMN: str})


def build_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
    return buf.getvalue()


def build_json_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    records = []
    for sheet_name, df in sheets.items():
        for row in df.to_dict(orient="records"):
            row["_sheet"] = sheet_name
            records.append(row)
    return json.dumps(records, ensure_ascii=False, indent=2, default=str).encode("utf-8")


def matches_categories(key: str, category_slugs: list[str]) -> bool:
    """True if any of the given slugs appears in the R2 key (case-insensitive,
    '_' and '-' treated the same)."""
    if not category_slugs:
        return True
    normalized_key = key.lower().replace("_", "-")
    return any(slug in normalized_key for slug in category_slugs)


# --------------------------------------------------------------------------
# Agent / agency slug extraction
# --------------------------------------------------------------------------

def parse_dict_field(value: Any) -> dict:
    if isinstance(value, dict):
        return value
    if isinstance(value, str):
        s = value.strip()
        if not s or s.lower() in {"nan", "none", "null"}:
            return {}
        try:
            return json.loads(s)
        except Exception:
            pass
        try:
            obj = ast.literal_eval(s)
            return obj if isinstance(obj, dict) else {}
        except Exception:
            return {}
    return {}


def extract_agent_slug(row) -> tuple[str | None, str | None]:
    """A listing is posted either by an individual agent (row['agent_profile'])
    or an agency (row['agent']). Returns (profile_type, slug) or (None, None)."""
    agent_profile = parse_dict_field(row.get(AGENT_PROFILE_COLUMN))
    if agent_profile.get("slug"):
        return "agent", agent_profile["slug"]

    agent = parse_dict_field(row.get(AGENT_COLUMN))
    if agent.get("slug"):
        return "agency", agent["slug"]

    return None, None


def extract_en_url(raw: Any) -> str | None:
    if isinstance(raw, dict):
        return raw.get("en") or raw.get("ar")
    if not isinstance(raw, str):
        return None
    s = raw.strip()
    if not s:
        return None
    if s.startswith("{"):
        try:
            obj = ast.literal_eval(s)
            if isinstance(obj, dict):
                return obj.get("en") or obj.get("ar")
        except Exception:
            pass
    return s


# --------------------------------------------------------------------------
# Profile phone cache (analogous to motors' users-data.xlsx)
# --------------------------------------------------------------------------

def profile_cache_key(profile_type: str, slug: str) -> str:
    return f"{profile_type}:{slug}"


def read_cached_profiles(client, profiles_key: str) -> dict[str, dict]:
    try:
        data = download_bytes(client, profiles_key)
    except client.exceptions.NoSuchKey:
        return {}
    except Exception as exc:
        print(f"[WARN] Could not read {profiles_key}: {exc}")
        return {}

    sheets = excel_sheets(data)
    if not sheets:
        return {}

    profiles = {}
    for df in sheets.values():
        if "profile_type" not in df.columns or "slug" not in df.columns:
            continue
        for _, row in df.iterrows():
            profile_type = row.get("profile_type")
            slug = row.get("slug")
            if is_empty(profile_type) or is_empty(slug):
                continue
            phone = row.get(PHONE_COLUMN)
            if not is_empty(phone):
                profiles[profile_cache_key(profile_type, slug)] = row.to_dict()
    return profiles


# --------------------------------------------------------------------------
# Challenge detection (shared)
# --------------------------------------------------------------------------

CHALLENGE_MARKERS = [
    "Pardon Our Interruption",
    "Additional security check is required",
    "I am human",
    "hCaptcha",
    "reeseSkipExpirationCheck",
]


def is_challenge_page(html: str) -> bool:
    return any(x in html for x in CHALLENGE_MARKERS)


def safe_content(page, retries=3, delay=1500) -> str:
    for attempt in range(retries):
        try:
            return page.content()
        except Exception:
            if attempt == retries - 1:
                raise
            page.wait_for_timeout(delay)
    return ""


# --------------------------------------------------------------------------
# Description scraping (listing page)
# --------------------------------------------------------------------------

def extract_description(page) -> str | None:
    try:
        see_more = page.locator('button:has-text("See full description")').first
        if see_more.is_visible(timeout=2000):
            see_more.click()
            page.wait_for_timeout(1500)
    except Exception:
        pass

    selectors = [
        'div[data-testid="description"]',
        '[data-testid="description"] + div',
        '[data-testid="description"] ~ div',
        '[data-testid="description-heading"]',
    ]
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=3000):
                text = loc.inner_text()
                if text and text.strip() and text.strip().lower() != "description":
                    return clean_excel_value(text.strip())
        except Exception:
            continue

    try:
        heading = page.locator('[data-testid="description"]').first
        if heading.is_visible(timeout=2000):
            text = heading.inner_text()
            if text and text.strip().lower() == "description":
                parent = page.locator('xpath=//*[@data-testid="description"]/..')
                divs = parent.locator("div").all()
                best = None
                for div in divs:
                    try:
                        t = div.inner_text()
                        if t and len(t.strip()) > len(best or ""):
                            best = t
                    except Exception:
                        pass
                if best and best.strip().lower() != "description":
                    return clean_excel_value(best.strip())
            elif text and text.strip():
                return clean_excel_value(text.strip())
    except Exception:
        pass

    return None


# --------------------------------------------------------------------------
# Phone scraping (agent / agency profile page)
# --------------------------------------------------------------------------

PROFILE_BUTTON_SELECTORS = [
    '[data-testid="profile-call-button"]',
    'button:has-text("Call")',
    '[data-testid="call-cta-button"]',
    'button:has-text("Show Phone Number")',
    'button:has-text("Show Number")',
    '[data-testid*="phone" i]',
    '[data-testid*="call" i]',
]

PROFILE_PHONE_KEY_HINTS = ["phone", "didnumber"]


def find_phone_recursive_profile(obj: Any) -> str | None:
    """Same idea as the listing-page phone search, but also matches
    'didNumber' -- the agent/agency profile GraphQL response returns the
    real number under didNumber and leaves phoneNumber null."""
    if isinstance(obj, dict):
        for key, value in obj.items():
            key_lower = str(key).lower()
            if any(hint in key_lower for hint in PROFILE_PHONE_KEY_HINTS) and isinstance(value, (str, int)) and value:
                return str(value)
        for value in obj.values():
            found = find_phone_recursive_profile(value)
            if found:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = find_phone_recursive_profile(item)
            if found:
                return found
    return None


def reveal_phone_from_profile(page, timeout_ms=10000) -> tuple[str | None, str]:
    captured = {"data": None}

    def handle_response(response):
        if "graphql" not in response.url:
            return
        try:
            post_data = response.request.post_data or ""
        except Exception:
            post_data = ""
        if "phone" in post_data.lower() and response.status == 200:
            try:
                captured["data"] = response.json()
            except Exception:
                pass

    page.on("response", handle_response)
    button = None
    for selector in PROFILE_BUTTON_SELECTORS:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=3500):
                button = loc
                break
        except Exception:
            continue

    if button is None:
        page.remove_listener("response", handle_response)
        return None, "button_not_found"

    try:
        button.scroll_into_view_if_needed()
        page.wait_for_timeout(300)
        try:
            button.click(timeout=6000)
        except Exception:
            button.click(force=True)

        waited = 0
        while captured["data"] is None and waited < timeout_ms:
            page.wait_for_timeout(400)
            waited += 400
    except Exception as exc:
        page.remove_listener("response", handle_response)
        return None, f"click_error: {exc}"
    finally:
        page.remove_listener("response", handle_response)

    if captured["data"] is None:
        return None, "no_response_captured"

    phone = find_phone_recursive_profile(captured["data"])
    if phone is None:
        return None, "phone_field_not_found"
    return phone, "ok"


# --------------------------------------------------------------------------
# prepare
# --------------------------------------------------------------------------

def prepare(date_str: str | None, out_dir: str, categories: str = "", job_type: str = "all"):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    jobs_dir = out / "jobs"
    jobs_dir.mkdir(exist_ok=True)

    date_iso, day_prefix = yesterday_prefix(date_str)
    client = r2_client()

    property_prefix = f"{day_prefix}{PROPERTY_ROOT}/"
    keys = list_keys(client, property_prefix)

    category_slugs = [c.strip().lower().replace("_", "-") for c in categories.split(",") if c.strip()]
    if category_slugs:
        before = len(keys)
        filtered = [k for k in keys if matches_categories(k, category_slugs)]
        if not filtered:
            print(f"[PREPARE][WARN] No files matched categories {category_slugs}.")
            print("[PREPARE][WARN] Available files were:")
            for k in keys:
                print(f"    {k}")
        keys = filtered
        print(f"[PREPARE] Category filter: {category_slugs}")
        print(f"[PREPARE] Files after category filter: {len(keys)} (was {before})")

    print(f"[PREPARE] Date: {date_iso}")
    print(f"[PREPARE] Found {len(keys)} property Excel file(s).")

    profiles_key = f"{property_prefix}profiles-data/profiles-data.xlsx"
    cached_profiles = read_cached_profiles(client, profiles_key)
    print(f"[PREPARE] Profiles cache: {profiles_key}")
    print(f"[PREPARE] Cached profiles with phone: {len(cached_profiles)}")

    description_work = []
    profiles_to_scrape: dict[str, dict] = {}

    for key in keys:
        sheets = excel_sheets(download_bytes(client, key))
        for sheet_name, df in sheets.items():
            if "id" not in df.columns:
                continue

            for row_pos, (_, row) in enumerate(df.iterrows()):
                listing_id = row.get("id")
                if is_empty(listing_id):
                    continue
                listing_id = str(listing_id)

                description_missing = is_empty(row.get(DESCRIPTION_COLUMN))
                if description_missing:
                    description_work.append({
                        "kind": "description",
                        "file_key": key,
                        "sheet_name": sheet_name,
                        "row_position": row_pos,
                        "id": listing_id,
                        "absolute_url": row.get(ABS_URL_COLUMN),
                    })

                phone_missing = is_empty(row.get(PHONE_COLUMN))
                if phone_missing:
                    profile_type, slug = extract_agent_slug(row)
                    if profile_type and slug:
                        cache_key = profile_cache_key(profile_type, slug)
                        if cache_key not in cached_profiles and cache_key not in profiles_to_scrape:
                            profiles_to_scrape[cache_key] = {
                                "kind": "profile",
                                "profile_type": profile_type,
                                "slug": slug,
                            }

    print(f"[PREPARE] Total listings needing description: {len(description_work)}")
    print(f"[PREPARE] Total unique agent/agency profiles needing phone: {len(profiles_to_scrape)}")

    manifest = []
    description_chunks = []
    profile_chunks = []

    if job_type in ("all", "description"):
        description_chunks = [description_work[i:i + 15] for i in range(0, len(description_work), 15)]
        for idx, chunk in enumerate(description_chunks):
            path = jobs_dir / f"desc_{idx:05d}.json"
            path.write_text(json.dumps(chunk, ensure_ascii=False, default=str, indent=2), encoding="utf-8")
            manifest.append(str(path.relative_to(out)))

    if job_type in ("all", "profile"):
        profile_chunks = [list(profiles_to_scrape.values())[i:i + 10] for i in range(0, len(profiles_to_scrape), 10)]
        for idx, chunk in enumerate(profile_chunks):
            path = jobs_dir / f"profile_{idx:05d}.json"
            path.write_text(json.dumps(chunk, ensure_ascii=False, default=str, indent=2), encoding="utf-8")
            manifest.append(str(path.relative_to(out)))

    (out / "manifest.json").write_text(
        json.dumps({
            "date": date_iso,
            "prefix": day_prefix,
            "property_prefix": property_prefix,
            "category_slugs": category_slugs,
            "profiles_key": profiles_key,
            "listing_keys": keys,
            "jobs": manifest,
            "total_description_items": len(description_work) if job_type in ("all", "description") else 0,
            "total_profile_items": len(profiles_to_scrape) if job_type in ("all", "profile") else 0,
            "total_jobs": len(manifest),
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"[PREPARE] Jobs Generated: {len(manifest)} (job_type='{job_type}')")
    print(f"[PREPARE] Manifest: {out / 'manifest.json'}")


# --------------------------------------------------------------------------
# scrape
# --------------------------------------------------------------------------

def scrape_job(job_file: str, output_file: str):
    from camoufox.sync_api import Camoufox

    items = json.loads(Path(job_file).read_text(encoding="utf-8"))
    results = []

    with Camoufox(headless=True, humanize=True, geoip=True, block_images=False) as browser:
        page = browser.new_page()

        for n, item in enumerate(items, 1):
            kind = item.get("kind")
            print(f"\n[{n}/{len(items)}] kind={kind}")

            if kind == "description":
                result = _scrape_description_item(page, item)
            elif kind == "profile":
                result = _scrape_profile_item(page, item)
            else:
                print(f"  [WARN] Unknown kind: {kind}")
                continue

            results.append(result)

        page.close()

    Path(output_file).write_text(json.dumps(results, ensure_ascii=False, default=str, indent=2), encoding="utf-8")
    print(f"\nWrote {len(results)} result(s) to {output_file}")


def _scrape_description_item(page, item: dict) -> dict:
    listing_id = item["id"]
    result = {
        "kind": "description",
        "file_key": item["file_key"],
        "sheet_name": item["sheet_name"],
        "row_position": item["row_position"],
        "id": listing_id,
        "description_full": None,
        "description_status": None,
    }

    url = extract_en_url(item.get("absolute_url"))
    if not url:
        print("  [ERROR] No URL")
        result["description_status"] = "no_url"
        return result

    print(f"  [URL] {url}")

    for attempt in range(2):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(random.uniform(6000, 10000))

            html = safe_content(page)
            if is_challenge_page(html):
                print("  [CHALLENGE] Imperva challenge detected")
                result["description_status"] = "imperva_challenge"
                return result

            description = extract_description(page)
            result["description_full"] = description
            result["description_status"] = "ok" if description else "not_found"
            print(f"  -> description: {'ok' if description else 'not found'}")
            return result
        except Exception as exc:
            if attempt == 0:
                print(f"  -> Transient error, retrying once: {exc}")
                page.wait_for_timeout(2000)
                continue
            result["description_status"] = f"error: {exc}"
            print(f"  -> FAILED after retry: {exc}")
            return result

    return result


def _scrape_profile_item(page, item: dict) -> dict:
    profile_type = item["profile_type"]
    slug = item["slug"]
    result = {
        "kind": "profile",
        "profile_type": profile_type,
        "slug": slug,
        "phone": None,
        "phone_status": None,
    }

    base_url = PROFILE_BASE_URLS.get(profile_type)
    if not base_url:
        result["phone_status"] = f"unknown_profile_type: {profile_type}"
        return result

    url = base_url.format(slug=slug)
    print(f"  [URL] {url}")

    for attempt in range(2):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(random.uniform(10000, 15000))

            html = safe_content(page)
            if is_challenge_page(html):
                print("  [CHALLENGE] Imperva challenge detected")
                result["phone_status"] = "imperva_challenge"
                return result

            phone, status = reveal_phone_from_profile(page)
            result["phone"] = phone
            result["phone_status"] = status
            print(f"  -> phone: {phone} (status: {status})")
            return result
        except Exception as exc:
            if attempt == 0:
                print(f"  -> Transient error, retrying once: {exc}")
                page.wait_for_timeout(2000)
                continue
            result["phone_status"] = f"error: {exc}"
            print(f"  -> FAILED after retry: {exc}")
            return result

    return result


# --------------------------------------------------------------------------
# combine-profiles (NEW)
# --------------------------------------------------------------------------

def combine_profiles(results_dir: str, date_str: str | None):
    root = Path(results_dir)
    manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
    profiles_key = manifest["profiles_key"]
    client = r2_client()

    result_files = sorted((root / "results").glob("*.json"))
    all_results = []
    for path in result_files:
        try:
            all_results.extend(json.loads(path.read_text(encoding="utf-8")))
        except Exception as e:
            print(f"[WARN] Failed to read result file {path}: {e}")

    new_profiles = {}
    for r in all_results:
        if r.get("kind") == "profile":
            profile_type = r.get("profile_type")
            slug = r.get("slug")
            phone = r.get("phone")
            if profile_type and slug and not is_empty(phone) and r.get("phone_status") == "ok":
                new_profiles[profile_cache_key(profile_type, slug)] = {
                    "profile_type": profile_type,
                    "slug": slug,
                    PHONE_COLUMN: phone,
                    "updated_at": datetime.now(timezone.utc).isoformat(),
                }

    print(f"[COMBINE-PROFILES] Found {len(new_profiles)} new profile phone numbers.")

    old_profiles = read_cached_profiles(client, profiles_key)
    merged_profiles = {}
    for cache_key, row in old_profiles.items():
        merged_profiles[cache_key] = {
            "profile_type": row.get("profile_type"),
            "slug": row.get("slug"),
            PHONE_COLUMN: row.get(PHONE_COLUMN),
            "updated_at": row.get("updated_at"),
        }
    for cache_key, row in new_profiles.items():
        if cache_key not in merged_profiles or is_empty(merged_profiles[cache_key].get(PHONE_COLUMN)):
            merged_profiles[cache_key] = row

    profiles_df = pd.DataFrame(list(merged_profiles.values()))
    if not profiles_df.empty:
        profiles_df = profiles_df.drop_duplicates(subset=["profile_type", "slug"], keep="last")
        upload_bytes(
            client,
            profiles_key,
            build_excel_bytes({"profiles": profiles_df}),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        print(f"[COMBINE-PROFILES] Successfully uploaded updated profiles-data.xlsx to R2: {len(profiles_df)} total profiles")


# --------------------------------------------------------------------------
# combine
# --------------------------------------------------------------------------

def combine(results_dir: str, date_str: str | None):
    root = Path(results_dir)
    manifest = json.loads((root / "manifest.json").read_text(encoding="utf-8"))
    date_iso = manifest["date"]
    profiles_key = manifest["profiles_key"]
    listing_keys = manifest["listing_keys"]
    client = r2_client()

    result_files = sorted((root / "results").glob("*.json"))
    all_results = []
    for path in result_files:
        try:
            all_results.extend(json.loads(path.read_text(encoding="utf-8")))
        except Exception as e:
            print(f"[WARN] Failed to read result file {path}: {e}")

    description_updates = {}
    new_profiles = {}

    for r in all_results:
        if r.get("kind") == "description":
            key = (r["file_key"], r["sheet_name"], str(r["id"]))
            description_updates[key] = r
        elif r.get("kind") == "profile":
            profile_type = r.get("profile_type")
            slug = r.get("slug")
            phone = r.get("phone")
            if profile_type and slug and not is_empty(phone) and r.get("phone_status") == "ok":
                new_profiles[profile_cache_key(profile_type, slug)] = {
                    "profile_type": profile_type,
                    "slug": slug,
                    PHONE_COLUMN: phone,
                    "updated_at": datetime.now(timezone.utc).isoformat(),
                }

    # Merge profiles cache (new phones only added, cache never loses a phone
    # it already had).
    old_profiles = read_cached_profiles(client, profiles_key)
    merged_profiles = {}
    for cache_key, row in old_profiles.items():
        merged_profiles[cache_key] = {
            "profile_type": row.get("profile_type"),
            "slug": row.get("slug"),
            PHONE_COLUMN: row.get(PHONE_COLUMN),
            "updated_at": row.get("updated_at"),
        }
    for cache_key, row in new_profiles.items():
        if cache_key not in merged_profiles or is_empty(merged_profiles[cache_key].get(PHONE_COLUMN)):
            merged_profiles[cache_key] = row

    profiles_df = pd.DataFrame(list(merged_profiles.values()))
    if not profiles_df.empty:
        profiles_df = profiles_df.drop_duplicates(subset=["profile_type", "slug"], keep="last")
        upload_bytes(
            client,
            profiles_key,
            build_excel_bytes({"profiles": profiles_df}),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        print(f"[COMBINE] profiles-data: {len(profiles_df)} profiles")

    # Re-read EVERY property file for the day (not just ones touched by a
    # scrape job) and fill in description_full + contact_phone_number.
    changed_files = 0
    phones_filled = 0
    descriptions_filled = 0

    for key in listing_keys:
        try:
            sheets = excel_sheets(download_bytes(client, key))
        except Exception as exc:
            print(f"[WARN] Could not read {key}: {exc}")
            continue

        changed = False

        for sheet_name, df in sheets.items():
            if "id" not in df.columns:
                continue
            if PHONE_COLUMN not in df.columns:
                df[PHONE_COLUMN] = None
            df[PHONE_COLUMN] = df[PHONE_COLUMN].astype(object)
            if DESCRIPTION_COLUMN not in df.columns:
                df[DESCRIPTION_COLUMN] = None
            df[DESCRIPTION_COLUMN] = df[DESCRIPTION_COLUMN].astype(object)

            for pos, (_, row) in enumerate(df.iterrows()):
                listing_id = row.get("id")
                if is_empty(listing_id):
                    continue

                u = description_updates.get((key, sheet_name, str(listing_id)))
                if u and u.get("description_status") == "ok" and not is_empty(u.get("description_full")):
                    df.at[df.index[pos], DESCRIPTION_COLUMN] = u["description_full"]
                    changed = True
                    descriptions_filled += 1

                if is_empty(row.get(PHONE_COLUMN)):
                    profile_type, slug = extract_agent_slug(row)
                    if profile_type and slug:
                        cached = merged_profiles.get(profile_cache_key(profile_type, slug))
                        phone = cached.get(PHONE_COLUMN) if cached else None
                        if not is_empty(phone):
                            df.at[df.index[pos], PHONE_COLUMN] = str(phone).strip()
                            changed = True
                            phones_filled += 1

            sheets[sheet_name] = df

        if changed:
            upload_bytes(
                client,
                key,
                build_excel_bytes(sheets),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            json_key = key[:-5] + ".json"
            upload_bytes(client, json_key, build_json_bytes(sheets), "application/json")
            changed_files += 1
            print(f"[COMBINE] Uploaded: {key}")

    summary = {
        "date": date_iso,
        "result_files": len(result_files),
        "result_rows": len(all_results),
        "listing_files_scanned": len(listing_keys),
        "changed_listing_files": changed_files,
        "descriptions_filled": descriptions_filled,
        "phones_filled": phones_filled,
        "cached_profiles_before": len(old_profiles),
        "successful_profiles_in_results": len(new_profiles),
        "cached_profiles_after": len(merged_profiles),
    }
    (root / "combine_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


# --------------------------------------------------------------------------
# CLI Entrypoint
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Property Enrichment Script")
    subparsers = parser.add_subparsers(dest="command", required=True)

    prepare_parser = subparsers.add_parser("prepare")
    prepare_parser.add_argument("--date", type=str, default=None)
    prepare_parser.add_argument("--out", type=str, default="work")
    prepare_parser.add_argument("--categories", type=str, default="")
    prepare_parser.add_argument(
        "--job-type", 
        type=str, 
        choices=["all", "profile", "description"], 
        default="all",
        help="Filter jobs by type: profile, description, or all"
    )

    scrape_parser = subparsers.add_parser("scrape")
    scrape_parser.add_argument("--job", type=str, required=True)
    scrape_parser.add_argument("--output", type=str, required=True)

    combine_profiles_parser = subparsers.add_parser("combine-profiles")
    combine_profiles_parser.add_argument("--work", type=str, required=True)
    combine_profiles_parser.add_argument("--date", type=str, default=None)

    combine_parser = subparsers.add_parser("combine")
    combine_parser.add_argument("--work", type=str, required=True)
    combine_parser.add_argument("--date", type=str, default=None)

    args = parser.parse_args()

    if args.command == "prepare":
        prepare(date_str=args.date, out_dir=args.out, categories=args.categories or "", job_type=args.job_type)
    elif args.command == "scrape":
        scrape_job(job_file=args.job, output_file=args.output)
    elif args.command == "combine-profiles":
        combine_profiles(results_dir=args.work, date_str=args.date)
    elif args.command == "combine":
        combine(results_dir=args.work, date_str=args.date)


if __name__ == "__main__":
    main()