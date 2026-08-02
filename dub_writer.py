import pandas as pd
import json
import ast
import os
import re
import io
import random
import time
import requests as req
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed
from r2_uploader import upload_buffer
from datetime import datetime, timedelta, timezone

from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth


PROPERTY_CATEGORIES = {
    "rent_residential", "rent_commercial", "rent_rooms_rent_flatmates",
    "rent_holiday_homes", "rent_short_term_monthly", "rent_short_term_daily",
    "sale_residential", "sale_commercial", "sale_land", "sale_multiple_units"
}

JOB_CATEGORIES = {"jobs", "jobs_wanted"}
NO_IMAGE_CATEGORIES = {"jobs", "jobs_wanted"}

COLUMNS_TO_DROP = ["tag_slugs", "category", "category_slug_tree", "category_tree", "categories_v2",
                  "site_categories_slug_tree", "permalink", "short_url", "short_url_v2",
                  "rent_is_paid", "_highlightResult", "categories", "photo"]

# Long-edge cap (px) images are downscaled to before upload, plus the WEBP
# quality used when re-encoding. Most source photos are 3000px+ wide;
# capping the long edge is what actually cuts stored bytes -- quality alone
# only goes so far.
MAX_IMAGE_DIMENSION = 1280
WEBP_QUALITY = 65


def parse_dict_field(value):
    if isinstance(value, dict):
        return value
    if isinstance(value, str):
        try:
            return json.loads(value)
        except (json.JSONDecodeError, TypeError):
            try:
                return ast.literal_eval(value)
            except Exception:
                return {}
    return {}


def get_category_names(category_v2_value) -> list:
    cat = parse_dict_field(category_v2_value)
    return cat.get("names_en", [])


def sanitize_name(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", str(name))
    name = name.replace(" ", "_")
    name = re.sub(r'_+', '_', name)
    return name.strip("_")


def build_property_meta(names_en: list) -> dict:
    """
    category_v2.names_en for property listings is ordered leaf -> top,
    e.g. ['Commercial Building', 'Commercial', 'Property for Sale'].

    cat0     = top-level folder            (names_en[-1], e.g. "Property for Sale")
    filename = the {filename}.xlsx/json    (names_en[-2], e.g. "Commercial";
               falls back to names_en[0] itself when there's no level above it,
               e.g. ['Land', 'Property for Sale'] -> filename "Land")
    sheet    = sheet inside that file       (names_en[0], the deepest leaf,
               e.g. "Commercial Building"; equals filename itself when there's
               no deeper level, so a single-sheet file still gets one sheet
               named after itself instead of "Other")
    """
    if not names_en:
        return {"cat0": "Property", "filename": "Other", "sheet": "Other"}

    top = names_en[-1]
    filename_source = names_en[-2] if len(names_en) >= 2 else names_en[0]
    sheet_source = names_en[0] if len(names_en) >= 3 else filename_source

    return {"cat0": top, "filename": filename_source, "sheet": sheet_source}


def build_job_meta(names_en: list, category_name: str) -> dict:
    """
    category_v2.names_en for job listings is ordered top -> leaf,
    e.g. ['Jobs Wanted', 'Engineering', 'Civil Engineering'].

    cat0     = top-level folder         (names_en[0], e.g. "Jobs Wanted"/"Jobs")
    filename = the {filename}.xlsx/json (names_en[1], e.g. "Engineering";
               falls back to names_en[0] when there's no level below it)
    sheet    = sheet inside that file   (names_en[2], the deepest leaf,
               e.g. "Civil Engineering"; equals filename itself when
               there's no deeper level, e.g. ['Jobs', 'Manufacturing / Warehouse'])
    """
    if not names_en:
        return {"cat0": category_name, "filename": "Other", "sheet": "Other"}

    top = names_en[0]
    filename_source = names_en[1] if len(names_en) >= 2 else names_en[0]
    sheet_source = names_en[2] if len(names_en) >= 3 else filename_source

    return {"cat0": top, "filename": filename_source, "sheet": sheet_source}


def extract_image_urls(row: pd.Series) -> list:
    if "photo_mains" in row and isinstance(row["photo_mains"], list):
        return row["photo_mains"]

    if "photos" in row and isinstance(row["photos"], list):
        urls = []
        for item in row["photos"]:
            if isinstance(item, dict) and item.get("main"):
                urls.append(item["main"])
        return urls

    return []


AGENT_PROFILE_FIELDS = [
    ("agent_profile", "agent"),   # individual agent -> property-agents/{slug}/
    ("agent", "agency"),          # agency -> property-agencies/{slug}/
]


def extract_agent_slug(row: pd.Series):
    """
    A listing is posted either by an individual agent (row['agent_profile'])
    or an agency (row['agent']). Returns (profile_type, slug) for whichever
    is present, or (None, None) if neither field has a usable slug.
    """
    for field_name, profile_type in AGENT_PROFILE_FIELDS:
        value = row.get(field_name)
        parsed = parse_dict_field(value) if not isinstance(value, dict) else value
        if isinstance(parsed, dict):
            slug = parsed.get("slug")
            if slug:
                return profile_type, slug
    return None, None


def generate_data_quality_report(df: pd.DataFrame, total_rows: int) -> str:
    report_lines = ["--- Data Quality Report ---"]
    for col in df.columns:
        missing = df[col].isna().sum() + (df[col] == '').sum()
        pct = (missing / total_rows) * 100 if total_rows > 0 else 0
        report_lines.append(f'  {col}: {missing} empty ({pct:.2f}%)')
    return "\n".join(report_lines)


def load_all_hits(jsonl_files: list) -> pd.DataFrame:
    rows = []
    for path in jsonl_files:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    rows.append(json.loads(line))
    df = pd.DataFrame(rows)

    existing_cols = [c for c in COLUMNS_TO_DROP if c in df.columns]
    if existing_cols:
        df = df.drop(columns=existing_cols)
        print(f"  Dropped columns: {existing_cols}")

    return df

OFF_PLAN_SOURCE_CATEGORY = "sale_residential"
OFF_PLAN_STATUS_VALUE = "off_plan"


def split_off_plan(df: pd.DataFrame) -> dict:
    if "completion_status" not in df.columns:
        print(f"  \u26a0\ufe0f Column 'completion_status' not found, skipping off_plan split.")
        return {"sale_residential": df}

    is_off_plan = df["completion_status"] == OFF_PLAN_STATUS_VALUE

    off_plan_df = df[is_off_plan].copy()
    rest_df = df[~is_off_plan].copy()

    print(f"  Split sale_residential: off_plan={len(off_plan_df)}, rest={len(rest_df)}")

    return {
        "off_plan": off_plan_df,
        "sale_residential": rest_df,
    }

def _get_english_url(absolute_url_value):
    parsed = parse_dict_field(absolute_url_value)
    if isinstance(parsed, dict):
        return parsed.get("en") or parsed.get("ar")
    if isinstance(absolute_url_value, str):
        return absolute_url_value
    return None

DESCRIPTION_SELECTORS = [
    '[data-testid="description"]',
    '[data-testid="description-heading"]',
]

def _extract_description(page):
    for selector in DESCRIPTION_SELECTORS:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=3000):
                text = loc.inner_text()
                if text:
                    return text
        except Exception:
            continue
    return None

def enrich_with_description(
    df: pd.DataFrame,
    url_column: str = "absolute_url",
    headless: bool = True,
    min_delay: float = 5,
    max_delay: float = 12,
) -> pd.DataFrame:
    df = df.copy()
    description_col = [None] * len(df)

    with Stealth().use_sync(sync_playwright()) as p:
        browser = p.chromium.launch(headless=headless, channel="chrome")
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            locale="en-US",
            timezone_id="Asia/Dubai",
        )
        page = context.new_page()

        for pos, (idx, row) in enumerate(df.iterrows()):
            url = _get_english_url(row.get(url_column))
            if not url:
                print(f"  [{pos + 1}/{len(df)}] Skipped - no URL")
                continue

            print(f"  [{pos + 1}/{len(df)}] Visiting: {url}")
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(random.uniform(1500, 3000))

                html = page.content()
                if "Pardon Our Interruption" in html:
                    print("    -> Imperva challenge hit, stopping enrichment.")
                    break

                description_col[pos] = _extract_description(page)

            except Exception as e:
                print(f"    -> FAILED: {e}")

            if pos < len(df) - 1:
                delay = random.uniform(min_delay, max_delay)
                time.sleep(delay)

        page.close()
        browser.close()
    df["description_full"] = description_col
    return df


def download_images(images: list, id_prod: str = "", category: str = "", cat0: str = "") -> list:
    r2_paths = []
    uploaded = 0
    failed = 0

    if not images or not isinstance(images, list):
        return r2_paths

    ext = "webp"
    file_prefix = id_prod or "unknown"
    category_display = cat0

    for idx, img_url in enumerate(images, start=1):
        filename = f"{file_prefix}-{idx}.{ext}"
        try:
            r = req.get(img_url, timeout=15)
            if r.status_code == 200:
                img = Image.open(io.BytesIO(r.content))
                img = img.convert("RGB")
                #img.thumbnail((MAX_IMAGE_DIMENSION, MAX_IMAGE_DIMENSION), Image.LANCZOS)
                output_buffer = io.BytesIO()
                img.save(output_buffer, format="WEBP", quality=WEBP_QUALITY, method=6)
                output_buffer.seek(0)

                r2_key = upload_buffer(
                    output_buffer,
                    filename=filename,
                    folder_name="DUAE",
                    category=category,
                    file_type="images",
                    content_type="image/webp",
                    dt=None,
                    category_display=category_display
                )
                if r2_key:
                    r2_paths.append(r2_key)
                    uploaded += 1
                else:
                    failed += 1
            else:
                failed += 1
        except Exception as e:
            print(f"    [ERROR] {filename} image {idx}: {e}")
            failed += 1

    if uploaded or failed:
        print(f"    {file_prefix}: {uploaded} uploaded, {failed} failed out of {len(images)}")
    return r2_paths


def process_images_for_group(df: pd.DataFrame, category: str, cat0: str, workers: int = 4) -> pd.DataFrame:
    df = df.copy()
    n = len(df)
    results = [None] * n

    def worker(pos: int, images: list, id_prod: str) -> tuple:
        r2_paths = download_images(images, id_prod=id_prod, category=category, cat0=cat0)
        return pos, r2_paths

    tasks = []
    for pos, (idx, row) in enumerate(df.iterrows()):
        images = extract_image_urls(row)
        id_prod = str(row.get("id", idx))
        tasks.append((pos, images, id_prod))

    print(f"  Downloading images for {n} listings using {workers} workers...")

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(worker, pos, images, id_prod): pos for pos, images, id_prod in tasks}
        completed = 0
        for future in as_completed(futures):
            try:
                pos, r2_paths = future.result(timeout=120)
                results[pos] = r2_paths
            except Exception as e:
                pos = futures[future]
                print(f"    [ERROR] Task {pos} failed: {e}")
                results[pos] = []
            completed += 1
            if completed % 50 == 0 or completed == n:
                print(f"    Progress: {completed}/{n}")

    df["images_r2_paths"] = results
    return df


def _write_excel_and_json(sheets: dict, xlsx_path: str, json_path: str) -> tuple:
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    all_records = []
    for sheet_name, df in sheets.items():
        records = df.to_dict(orient="records")
        for r in records:
            r["_sheet"] = sheet_name
        all_records.extend(records)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_records, f, ensure_ascii=False, indent=2, default=str)

    return xlsx_path, json_path


def build_group_summary(file_groups: dict, group_df: pd.DataFrame, cat0: str, dt: datetime) -> dict:
    """
    file_groups: {filename -> full df for that file} -- one entry per
    {filename}.xlsx written under this cat0 (e.g. "Residential",
    "Commercial", "Land"), NOT per internal sheet.
    """
    subcategories = [
        {
            "name_ar": "",
            "name_en": name,
            "slug": name,
            "listings_count": len(fdf),
            "has_subcategories": False,
            "subcategories": [],
        }
        for name, fdf in file_groups.items()
    ]
    return {
        "scraped_at": dt.isoformat(),
        "data_scraped_date": (dt - timedelta(days=1)).strftime("%Y-%m-%d"),
        "saved_to_R2_date": dt.strftime("%Y-%m-%d"),
        "category": cat0,
        "total_subcategories": len(subcategories),
        "total_listings": len(group_df),
        "subcategories": subcategories,
    }
def convert_timestamp_columns(df: pd.DataFrame) -> pd.DataFrame:
    timestamp_columns = [
        "added",
        "created_at",
        "last_updated_at"
    ]
    df = df.copy()
    for col in timestamp_columns:
        if col in df.columns:
            df[col] = (
                pd.to_datetime(
                    pd.to_numeric(df[col], errors="coerce"),
                    unit="s",
                    errors="coerce",
                    utc=True
                )
                .dt.tz_convert("Asia/Dubai")
                .dt.strftime("%Y-%m-%d %H:%M:%S")
            )

            print(f"  Converted timestamp column: {col}")

    return df

def _process_category_internal(category_name: str, df: pd.DataFrame, output_base_dir: str,
                                 upload_images: bool, image_workers: int) -> dict:
    if df.empty:
        return {"excel_files": [], "json_files": []}

    df["_names_en"] = df["category_v2"].apply(get_category_names)

    if category_name in PROPERTY_CATEGORIES or category_name == "off_plan":
        meta_list = df["_names_en"].apply(build_property_meta)
    elif category_name in JOB_CATEGORIES:
        meta_list = df["_names_en"].apply(lambda n: build_job_meta(n, category_name))
    else:
        print(f"        Unknown category family for '{category_name}', skipping.")
        return {"excel_files": [], "json_files": []}

    df["_cat0"] = meta_list.apply(lambda m: m["cat0"])
    df["_filename"] = meta_list.apply(lambda m: m["filename"])
    df["_sheet"] = meta_list.apply(lambda m: m["sheet"])

    if "id" in df.columns:
        df = df.drop_duplicates(subset=["id"], keep="first")

    excel_files = []
    json_files = []

    has_image_column = "photo_mains" in df.columns or "photos" in df.columns
    should_process_images = upload_images and has_image_column and category_name not in NO_IMAGE_CATEGORIES

    cols_to_drop = ["_cat0", "_filename", "_sheet", "_names_en"]

    for cat0, group_df in df.groupby("_cat0", dropna=False):
        safe_cat0 = sanitize_name(cat0)

        group_quality_report = generate_data_quality_report(group_df, len(group_df))

        group_dir = os.path.join(output_base_dir, safe_cat0)
        os.makedirs(group_dir, exist_ok=True)

        if should_process_images:
            print(f"  Processing images for {safe_cat0} ({len(group_df)} listings)...")
            group_df = process_images_for_group(
                group_df, category=category_name, cat0=safe_cat0, workers=image_workers
            )

        excel_dir = os.path.join(group_dir, "excel")
        json_dir = os.path.join(group_dir, "json")
        summary_dir = os.path.join(group_dir, "summary")
        os.makedirs(excel_dir, exist_ok=True)
        os.makedirs(json_dir, exist_ok=True)
        os.makedirs(summary_dir, exist_ok=True)

        file_groups = {}  # filename -> full df, for summary.json counts

        for filename, f_df in group_df.groupby("_filename"):
            safe_filename = sanitize_name(filename)
            main_xlsx = os.path.join(excel_dir, f"{safe_filename}.xlsx")
            main_json = os.path.join(json_dir, f"{safe_filename}.json")

            sheets = {}
            for sheet_name, sdf in f_df.groupby("_sheet"):
                sdf_clean = sdf.drop(columns=[c for c in cols_to_drop if c in sdf.columns])
                safe_sheet = sanitize_name(sheet_name)[:31]
                sheets[safe_sheet] = sdf_clean

            xlsx_path, json_path = _write_excel_and_json(sheets, main_xlsx, main_json)
            excel_files.append(xlsx_path)
            json_files.append(json_path)
            print(f"  Saved: {main_xlsx} ({len(f_df)} rows, {len(sheets)} sheet(s))")

            file_groups[safe_filename] = f_df

        dt = datetime.now(timezone.utc)
        summary = build_group_summary(file_groups, group_df, safe_cat0, dt)
        summary_file_path = os.path.join(summary_dir, "summary.json")
        with open(summary_file_path, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"  Saved summary: {summary_file_path} ({summary['total_subcategories']} subcats, {summary['total_listings']} listings)")

    return {"excel_files": excel_files, "json_files": json_files}


def process_category(category_name: str, jsonl_files: list, output_base_dir: str,
                      upload_images: bool = True, image_workers: int = 4,
                      enrich_contact_details: bool = False,
                      phone_lookup: dict = None) -> dict:
    df = load_all_hits(jsonl_files)
    if df.empty:
        return {"total": 0, "excel_files": [], "json_files": []}

    df = convert_timestamp_columns(df)

    if enrich_contact_details and "absolute_url" in df.columns:
        print(f"  Enriching {len(df)} rows with description_full...")
        df = enrich_with_description(df)

    if phone_lookup:
        agent_info = df.apply(extract_agent_slug, axis=1, result_type="expand")
        agent_info.columns = ["_agent_profile_type", "_agent_slug"]
        df["contact_phone_number"] = agent_info["_agent_slug"].map(phone_lookup)
        matched = df["contact_phone_number"].notna().sum()
        print(f"  Matched phone numbers for {matched}/{len(df)} rows from phone_lookup")

    total = len(df)
    excel_files = []
    json_files = []

    if category_name == OFF_PLAN_SOURCE_CATEGORY:
        splits = split_off_plan(df)
        for split_name, split_df in splits.items():
            if split_df.empty:
                continue
            result = _process_category_internal(split_name, split_df, output_base_dir, upload_images, image_workers)
            excel_files.extend(result["excel_files"])
            json_files.extend(result["json_files"])
    else:
        result = _process_category_internal(category_name, df, output_base_dir, upload_images, image_workers)
        excel_files.extend(result["excel_files"])
        json_files.extend(result["json_files"])

    return {"total": total, "excel_files": excel_files, "json_files": json_files}