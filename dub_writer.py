import pandas as pd
import json
import ast
import os
import re
import io
import requests as req
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed
from r2_uploader import upload_buffer
from datetime import datetime

PROPERTY_CATEGORIES = {
    "rent_residential", "rent_commercial", "rent_rooms_rent_flatmates",
    "rent_holiday_homes", "rent_short_term_daily",
    "sale_residential", "sale_commercial"
}

JOB_CATEGORIES = {"jobs", "jobs_wanted"}
NO_IMAGE_CATEGORIES = {"jobs", "jobs_wanted"}


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


def get_city_name(site_value) -> str:
    site = parse_dict_field(site_value)
    if not site:
        return "Unknown"

    if "en" in site:
        return site.get("en", "Unknown")

    name_field = site.get("name")
    if isinstance(name_field, dict):
        return name_field.get("en", "Unknown")
    if isinstance(name_field, str):
        return name_field

    return "Unknown"


def get_category_names(category_v2_value) -> list:
    cat = parse_dict_field(category_v2_value)
    return cat.get("names_en", [])


def sanitize_name(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", str(name))
    name = name.replace(" ", "_")
    return name.strip()


def build_property_meta(names_en: list, category_name: str) -> dict:
    if len(names_en) < 2:
        return {"cat0": "Property", "cat1": "Other", "filename": category_name, "sheet": "Other"}

    leaf = names_en[0]
    mid = names_en[1]
    top = names_en[-1]

    sheet = f"{mid} ({leaf})"

    return {"cat0": "Property", "cat1": top, "filename": category_name, "sheet": sheet}


def build_job_meta(names_en: list, category_name: str) -> dict:
    if not names_en:
        return {"cat0": category_name, "cat1": None, "filename": category_name, "sheet": "Other"}

    top = names_en[0]

    if len(names_en) >= 3:
        sheet = f"{names_en[1]} ({names_en[2]})"
    elif len(names_en) == 2:
        sheet = names_en[1]
    else:
        sheet = "Other"

    return {"cat0": top, "cat1": None, "filename": top, "sheet": sheet}


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
    return pd.DataFrame(rows)


def download_images(images: list, id_prod: str = "", category: str = "",
                     city: str = "", cat0: str = "", cat1: str = None) -> list:
    r2_paths = []
    uploaded = 0
    failed = 0

    if not images or not isinstance(images, list):
        return r2_paths

    ext = "webp"
    file_prefix = id_prod or "unknown"
    category_display = f"{cat0}/{cat1}" if cat1 else cat0

    for idx, img_url in enumerate(images, start=1):
        filename = f"{file_prefix}-{idx}.{ext}"
        try:
            r = req.get(img_url, timeout=15)
            if r.status_code == 200:
                img = Image.open(io.BytesIO(r.content))
                output_buffer = io.BytesIO()
                img = img.convert("RGB")
                img.save(output_buffer, format="WEBP", quality=100, method=6)
                output_buffer.seek(0)

                r2_key = upload_buffer(
                    output_buffer,
                    filename=filename,
                    folder_name="DUAE",
                    category=category,
                    file_type="images",
                    content_type="image/webp",
                    dt=None,
                    city=city,
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


def process_images_for_group(df: pd.DataFrame, category: str, city: str, cat0: str, cat1: str,
                              workers: int = 4) -> pd.DataFrame:
    df = df.copy()
    n = len(df)
    results = [None] * n

    def worker(pos: int, images: list, id_prod: str) -> tuple:
        r2_paths = download_images(images, id_prod=id_prod, category=category, city=city, cat0=cat0, cat1=cat1)
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


def process_category(category_name: str, jsonl_files: list, output_base_dir: str,
                      upload_images: bool = True, image_workers: int = 4,
                      city_filter: str = None) -> dict:
    df = load_all_hits(jsonl_files)
    if df.empty:
        return {"total": 0, "excel_files": [], "json_files": []}

    df["_city"] = df["site"].apply(get_city_name)

    if city_filter:
        df = df[df["_city"] == city_filter]
        print(f"  Filtered to city: {city_filter} ({len(df)} rows)")
        if df.empty:
            return {"total": 0, "excel_files": [], "json_files": []}

    df["_names_en"] = df["category_v2"].apply(get_category_names)

    if category_name in PROPERTY_CATEGORIES:
        meta_fn = build_property_meta
    elif category_name in JOB_CATEGORIES:
        meta_fn = build_job_meta
    else:
        print(f"  ⚠️ Unknown category family for '{category_name}', skipping.")
        return {"total": 0, "excel_files": [], "json_files": []}

    meta_list = df["_names_en"].apply(lambda n: meta_fn(n, category_name))
    df["_cat0"] = meta_list.apply(lambda m: m["cat0"])
    df["_cat1"] = meta_list.apply(lambda m: m["cat1"])
    df["_filename"] = meta_list.apply(lambda m: m["filename"])
    df["_sheet"] = meta_list.apply(lambda m: m["sheet"])

    if "id" in df.columns:
        df = df.drop_duplicates(subset=["id"], keep="first")

    excel_files = []
    json_files = []
    total = len(df)

    group_cols = ["_city", "_cat0", "_cat1", "_filename"]

    has_image_column = "photo_mains" in df.columns or "photos" in df.columns
    should_process_images = upload_images and has_image_column and category_name not in NO_IMAGE_CATEGORIES

    for keys, group_df in df.groupby(group_cols, dropna=False):
        city, cat0, cat1, filename = keys
        safe_city = sanitize_name(city)
        safe_cat0 = sanitize_name(cat0)
        safe_cat1 = sanitize_name(cat1) if pd.notna(cat1) and cat1 else None
        safe_filename = sanitize_name(filename)

        group_quality_report = generate_data_quality_report(group_df, len(group_df))

        if safe_cat1:
            group_dir = os.path.join(output_base_dir, safe_city, safe_cat0, safe_cat1)
        else:
            group_dir = os.path.join(output_base_dir, safe_city, safe_cat0)
        os.makedirs(group_dir, exist_ok=True)

        if should_process_images:
            print(f"  Processing images for {safe_city}/{safe_cat0}/{safe_cat1 or ''} ({len(group_df)} listings)...")
            group_df = process_images_for_group(
                group_df, category=category_name, city=safe_city,
                cat0=safe_cat0, cat1=safe_cat1, workers=image_workers
            )

        excel_dir = os.path.join(group_dir, "excel")
        json_dir = os.path.join(group_dir, "json")
        summary_dir = os.path.join(group_dir, "summary")
        os.makedirs(excel_dir, exist_ok=True)
        os.makedirs(json_dir, exist_ok=True)
        os.makedirs(summary_dir, exist_ok=True)

        main_xlsx = os.path.join(excel_dir, f"{safe_filename}.xlsx")
        main_json = os.path.join(json_dir, f"{safe_filename}.json")
        
        cols_to_drop = ["_city", "_cat0", "_cat1", "_filename", "_sheet", "_names_en"]
        sheets = {}
        for sheet_name, sdf in group_df.groupby("_sheet"):
            sdf_clean = sdf.drop(columns=[c for c in cols_to_drop if c in sdf.columns])
            safe_sheet = sanitize_name(sheet_name)[:31]
            sheets[safe_sheet] = sdf_clean

        xlsx_path, json_path = _write_excel_and_json(
            sheets,
            main_xlsx,
            main_json
        )
        excel_files.append(xlsx_path)
        json_files.append(json_path)
        print(f"  Saved: {main_xlsx} ({len(group_df)} rows)")

        summary_file_path = os.path.join(summary_dir, "summary.txt")
        with open(summary_file_path, "w", encoding="utf-8") as f:
            f.write(f"=== {category_name} ===\n")
            f.write(f"City: {safe_city}\n")
            f.write(f"Category: {safe_cat0}/{safe_cat1 or ''}\n")
            f.write(f"File: {safe_filename}\n")
            f.write(f"Total Rows: {len(group_df)}\n")
            f.write(f"Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(group_quality_report)

        print(f"  Saved summary: {summary_file_path}")

    return {"total": total, "excel_files": excel_files, "json_files": json_files}