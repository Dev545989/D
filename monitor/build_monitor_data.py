"""
build_monitor_data.py
=====================
Complete monitor data management.

================================================================================
COMPLETE USAGE GUIDE FOR build_monitor_data.py
================================================================================

COMMAND REFERENCE
────────────────────────────────────────────────────────────────────────────────
--prefix DOMAN              Required - specifies the site/R2 prefix
--upload-config            Upload local websites-config.yml to R2
--show-config              Display local config file content
--build-stats              Build monitor_stats.yml from Excel files
--days N                   Number of days to scan (default: 30, with --build-stats)
--no-save                  Don't upload stats to R2 (test mode only)
--output-stats FILE        Save stats to local YAML file
--report-date DATE         Generate report for specific date (YYYY-MM-DD)
--range-report             Generate aggregated report for a date range
--start-date DATE          Start of date range
--end-date DATE            End of date range
--output-report FILE       Save report to local JSON file

# =====================================================================
# build_monitor_data.py - Complete Usage Examples
# =====================================================================

# 1. UPLOAD CONFIG TO R2
# Uploads local websites-config.yml to R2 at DOMAN/monitor/websites-config.yml
python build_monitor_data.py --prefix DOMAN --upload-config

# 2. SHOW LOCAL CONFIG
# Prints the content of local websites-config.yml file
python build_monitor_data.py --prefix DOMAN --show-config

# 3. BUILD STATS (LAST 30 DAYS)
# Scans last 30 days of Excel files, builds monitor_stats.yml, uploads to R2
python build_monitor_data.py --prefix DOMAN --build-stats

# 4. BUILD STATS (CUSTOM DAYS)
# Scans only last 7 days (faster)
python build_monitor_data.py --prefix DOMAN --build-stats --days 7

# 5. BUILD STATS (LOCAL ONLY)
# Builds stats but saves locally as test_stats.yml (no R2 upload)
python build_monitor_data.py --prefix DOMAN --build-stats --days 7 --output-stats test_stats.yml

# 6. BUILD STATS (TEST MODE)
# Builds stats but doesn't upload to R2 (safe for testing)
python build_monitor_data.py --prefix DOMAN --build-stats --days 7 --no-save

# 7. GENERATE DAILY REPORT (YESTERDAY)
# Generates report for yesterday, saves to R2 at DOMAN/monitor/YYYY-MM-DD/report.json
python build_monitor_data.py --prefix DOMAN --report-date

# 8. GENERATE DAILY REPORT (SPECIFIC DATE)
# Generates report for July 28, 2026
python build_monitor_data.py --prefix DOMAN --report-date 2026-07-28

# 9. GENERATE REPORT (LOCAL ONLY)
# Generates report and saves locally as my_report.json
python build_monitor_data.py --prefix DOMAN --report-date 2026-07-28 --output-report my_report.json

# 10. GENERATE WEEKLY REPORT
# Aggregates reports from July 22 to July 28 into one report
python build_monitor_data.py --prefix DOMAN --range-report --start 2026-07-22 --end 2026-07-28

# 11. GENERATE MONTHLY REPORT
# Aggregates reports from July 1 to July 28
python build_monitor_data.py --prefix DOMAN --range-report --start 2026-07-01 --end 2026-07-28

# 12. GENERATE RANGE REPORT (LOCAL ONLY)
# Range report saved locally as monthly.json
python build_monitor_data.py --prefix DOMAN --range-report --start 2026-07-01 --end 2026-07-28 --output-report monthly.json

# 13. FULL RUN (STATS + REPORT)
# Builds stats for 30 days AND generates report for July 28
python build_monitor_data.py --prefix DOMAN --build-stats --days 30 --report-date 2026-07-28

# 14. FULL RUN (EVERYTHING LOCAL)
# Builds stats, generates report, saves everything locally
python build_monitor_data.py --prefix DOMAN --build-stats --days 30 --output-stats stats.yml --report-date 2026-07-28 --output-report report.json

# 15. DAILY AUTOMATION (GITHUB ACTIONS)
# Three commands: upload config, update stats, generate report
python build_monitor_data.py --prefix DOMAN --upload-config
python build_monitor_data.py --prefix DOMAN --build-stats --days 1
python build_monitor_data.py --prefix DOMAN --report-date

# 16. FIRST TIME SETUP
# Complete setup for a new site
python build_monitor_data.py --prefix DOMAN --upload-config
python build_monitor_data.py --prefix DOMAN --build-stats --days 30
python build_monitor_data.py --prefix DOMAN --report-date

# 17. CONFIG UPDATE
# After editing websites-config.yml locally
python build_monitor_data.py --prefix DOMAN --upload-config
python build_monitor_data.py --prefix DOMAN --build-stats --days 30

# 18. TEST NEW CONFIG (SAFE)
# Tests stats without affecting production R2 data
python build_monitor_data.py --prefix DOMAN --build-stats --days 1 --no-save --output-stats test_stats.yml

# 19. WEEKLY DEEP SCAN
# Full weekly scan for trend detection
python build_monitor_data.py --prefix DOMAN --build-stats --days 7

# 20. MONTHLY MAINTENANCE
# Full rebuild to catch any data gaps
python build_monitor_data.py --prefix DOMAN --build-stats --days 30

# =====================================================================
# FILE LOCATIONS IN R2
# =====================================================================
# Config:   DOMAN/monitor/websites-config.yml
# Stats:    DOMAN/monitor/monitor_stats.yml
# Report:   DOMAN/monitor/YYYY-MM-DD/report.json

# =====================================================================
# COMMON ERRORS & SOLUTIONS
# =====================================================================
# ERROR: "Config not found locally"
# SOLUTION: Create websites-config.yml in same folder as script

# ERROR: "No scrapers found in config"
# SOLUTION: Add scrapers section to websites-config.yml

# ERROR: "R2 connection failed"
# SOLUTION: Check AWS credentials in .env file

# ERROR: "No Excel files found"
# SOLUTION: Verify r2_path matches actual R2 structure
# =====================================================================

"""

import boto3
import json
import logging
import os
import json
import yaml
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any
from collections import defaultdict
from pathlib import Path

import openpyxl
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# Import from monitor_r2
from monitor_r2 import (
    build_r2_client,
    monitor_data_keys,
    report_r2_key,
    load_site_config_from_r2,
    fetch_yaml_object,
    put_yaml_object,
)

# Import existing functions
from inspect_r2_schema import (
    list_excel_files,
    download_excel,
    inspect_excel,
    accumulate_stats,
    r2_base_prefix,
    partition_date_for_data_date,
    excel_prefixes_for_date,
)

from ads_counter import (
    load_json_summaries,
    count_ads_from_downloads,
)

from r2_file_counter import count_site_r2_files

log = logging.getLogger("monitor")

# ── Local config path ──────────────────────────────────────────────────────
LOCAL_CONFIG_PATH = Path(__file__).parent / "websites-config.yml"


# ── Helper Functions ──────────────────────────────────────────────────────

def list_scraper_excel_files(client, bucket: str, r2_base: str, category: Optional[str], date_str: str) -> List[Dict]:
    """List all Excel files for a scraper's own category on a specific date."""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    part_dt = partition_date_for_data_date(dt)
    
    all_files = []
    seen_keys = set()
    
    for prefix in excel_prefixes_for_date(r2_base, category, part_dt):
        try:
            for f in list_excel_files(client, bucket, prefix):
                if f["key"] not in seen_keys:
                    seen_keys.add(f["key"])
                    f["date"] = dt.strftime("%Y-%m-%d")
                    all_files.append(f)
        except Exception as exc:
            log.warning(f"Error listing files in {prefix}: {type(exc).__name__}: {exc}")
    
    return all_files


def merge_stats(existing: Dict, new: Dict) -> Dict:
    """Merge new stats into existing stats structure."""
    if not new:
        return existing
    
    merged = dict(existing)
    
    if "observed_dates" in new:
        existing_dates = set(merged.get("observed_dates", []))
        existing_dates.update(new.get("observed_dates", []))
        merged["observed_dates"] = sorted(list(existing_dates))
    
    if "file_size_kb" in new:
        old_fs = merged.get("file_size_kb", {"min": None, "max": None, "last": None})
        new_fs = new.get("file_size_kb", {})
        merged["file_size_kb"] = {
            "min": min(
                old_fs.get("min") or float('inf'),
                new_fs.get("min") or float('inf')
            ) if old_fs.get("min") is not None or new_fs.get("min") is not None else None,
            "max": max(
                old_fs.get("max") or 0,
                new_fs.get("max") or 0
            ) if old_fs.get("max") is not None or new_fs.get("max") is not None else None,
            "last": new_fs.get("last") or old_fs.get("last")
        }
        if merged["file_size_kb"]["min"] == float('inf'):
            merged["file_size_kb"]["min"] = None
    
    if "sheets" in new:
        merged_sheets = dict(merged.get("sheets", {}))
        for sheet_name, new_sheet_data in new["sheets"].items():
            if sheet_name not in merged_sheets:
                merged_sheets[sheet_name] = new_sheet_data
            else:
                old_sheet = merged_sheets[sheet_name]
                new_row = new_sheet_data.get("row_count", {})
                old_row = old_sheet.get("row_count", {"min": None, "max": None, "last": None})
                merged_sheets[sheet_name]["row_count"] = {
                    "min": min(
                        old_row.get("min") or float('inf'),
                        new_row.get("min") or float('inf')
                    ) if old_row.get("min") is not None or new_row.get("min") is not None else None,
                    "max": max(
                        old_row.get("max") or 0,
                        new_row.get("max") or 0
                    ) if old_row.get("max") is not None or new_row.get("max") is not None else None,
                    "last": new_row.get("last") or old_row.get("last")
                }
                if merged_sheets[sheet_name]["row_count"]["min"] == float('inf'):
                    merged_sheets[sheet_name]["row_count"]["min"] = None
                
                old_cols = set(old_sheet.get("columns", []))
                new_cols = set(new_sheet_data.get("columns", []))
                merged_sheets[sheet_name]["columns"] = sorted(list(old_cols.union(new_cols)))
        
        merged["sheets"] = merged_sheets
    
    return merged


def get_stats_for_scraper(
    client,
    bucket: str,
    scraper_name: str,
    scraper_config: Dict,
    dates_to_check: List[str],
    existing_stats: Dict
) -> Dict:
    """Build stats for a single scraper by inspecting its Excel files."""
    r2_base, category = r2_base_prefix(scraper_config.get("r2_path", ""))
    if not r2_base:
        log.warning(f"  {scraper_name}: no r2_path in config — skipping")
        return {}
    
    stats = {}
    
    for date_str in dates_to_check:
        excel_files = list_scraper_excel_files(client, bucket, r2_base, category, date_str)
        
        if not excel_files:
            continue
        
        for xlsx_meta in excel_files:
            raw = download_excel(client, bucket, xlsx_meta["key"])
            if raw is None:
                continue
            
            inspected = inspect_excel(raw, xlsx_meta["key"])
            inspected["size_bytes"] = xlsx_meta["size_bytes"]
            
            accumulate_stats(
                stats,
                scraper_name,
                inspected,
                xlsx_meta["size_bytes"],
                date_str
            )
    
    if scraper_name in existing_stats:
        return merge_stats(existing_stats[scraper_name], stats)
    else:
        return stats


def get_categories_ads(
    client,
    bucket: str,
    scraper_configs: List[Dict],
    target_date: datetime,
    download_files: bool = True
) -> List[Dict[str, Any]]:
    """Get ads count per category/scraper for a specific date."""
    categories = []
    
    for scraper_config in scraper_configs:
        scraper_name = scraper_config.get("name")
        r2_base, category = r2_base_prefix(scraper_config.get("r2_path", ""))
        if not r2_base:
            continue
        
        part_dt = partition_date_for_data_date(target_date)
        
        json_total, json_key, json_breakdown = load_json_summaries(
            client, bucket, r2_base, part_dt, category=category
        )
        
        if json_total is not None and json_breakdown:
            for item in json_breakdown:
                if item.get("subcategory"):
                    categories.append({
                        "name": item["subcategory"],
                        "slug": item.get("level_3") or item["subcategory"].lower().replace(" ", "-"),
                        "total_ads": item.get("ads_count", 0),
                    })
            continue
        
        if download_files:
            all_xlsx = []
            seen_keys = set()
            
            for prefix in excel_prefixes_for_date(r2_base, category, part_dt):
                try:
                    for f in list_excel_files(client, bucket, prefix):
                        if f["key"] not in seen_keys:
                            seen_keys.add(f["key"])
                            all_xlsx.append(f)
                except Exception:
                    continue
            
            if not all_xlsx:
                continue
            
            downloads = []
            for xlsx_meta in all_xlsx:
                raw = download_excel(client, bucket, xlsx_meta["key"])
                if raw is not None:
                    downloads.append((xlsx_meta["key"], raw))
            
            if downloads:
                ads_stats = count_ads_from_downloads(downloads)
                unique_ads = ads_stats.get("unique_ads", 0)
                
                category_name = (
                    scraper_config.get("display_name") or 
                    scraper_config.get("category_name") or 
                    scraper_name
                )
                slug = scraper_config.get("slug") or scraper_name.lower().replace(" ", "-")
                
                categories.append({
                    "name": category_name,
                    "slug": slug,
                    "total_ads": unique_ads,
                })
    
    return categories


def collect_request_metrics(
    client,
    bucket: str,
    scraper_configs: List[Dict],
    target_date: datetime,
) -> Dict[str, Any]:
    """
    Collect request_metrics from all scrapers' summary.json for a given date.
    """
    metrics = {
        "requests_total": 0,
        "requests_failed": 0,
        "duration_sec": 0.0,
        "requests_per_min": 0.0,
        "error_rate_pct": None,
        "scrapers_with_metrics": 0,
    }
    
    for scraper_config in scraper_configs:
        scraper_name = scraper_config.get("name")
        r2_base, category = r2_base_prefix(scraper_config.get("r2_path", ""))
        if not r2_base:
            continue
        
        part_dt = partition_date_for_data_date(target_date)
        
        # Get summary.json from R2
        best_total, best_key, best_breakdown = load_json_summaries(
            client, bucket, r2_base, part_dt, category=category
        )
        
        if best_key:
            try:
                resp = client.get_object(Bucket=bucket, Key=best_key)
                data = json.loads(resp["Body"].read().decode("utf-8"))
                
                request_metrics = data.get("request_metrics", {})
                if request_metrics:
                    metrics["requests_total"] += request_metrics.get("requests_total", 0)
                    metrics["requests_failed"] += request_metrics.get("requests_failed", 0)
                    metrics["scrapers_with_metrics"] += 1
                    
                    duration = request_metrics.get("duration_sec", 0)
                    if duration and duration > metrics["duration_sec"]:
                        metrics["duration_sec"] = duration
            except Exception as e:
                pass
    
    if metrics["requests_total"] > 0:
        metrics["error_rate_pct"] = round(
            metrics["requests_failed"] / metrics["requests_total"] * 100.0, 2
        )
    
    if metrics["duration_sec"] > 0:
        metrics["requests_per_min"] = round(
            metrics["requests_total"] / (metrics["duration_sec"] / 60.0), 2
        )
    
    return metrics

# ── CONFIG FUNCTIONS ──────────────────────────────────────────────────────

def load_local_config() -> Dict:
    """Load config from local websites-config.yml file."""
    if not LOCAL_CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Local config not found at {LOCAL_CONFIG_PATH}. "
            f"Create websites-config.yml in the same directory as this script."
        )
    
    with open(LOCAL_CONFIG_PATH, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    
    log.info(f"Loaded local config from {LOCAL_CONFIG_PATH}")
    return config


def upload_config_to_r2(
    r2_prefix: str,
    config_data: Optional[Dict] = None,
    force: bool = True,
    client = None,
    bucket: Optional[str] = None,
) -> Dict:
    """
    Upload local websites-config.yml to R2.
    Always uploads from local file (default behavior).
    """
    if client is None or bucket is None:
        client, bucket = build_r2_client()
    
    if not bucket:
        raise ValueError("Bucket name is required")
    
    # Build site config for monitor_data_keys
    site = {
        "r2_prefix": r2_prefix,
        "folder": r2_prefix,
        "site_id": r2_prefix,
    }
    
    keys = monitor_data_keys(site)
    config_key = keys["config"]
    
    # Load from local file
    if config_data is None:
        try:
            config_data = load_local_config()
        except FileNotFoundError as e:
            return {
                "status": "error",
                "message": str(e),
                "path": config_key
            }
    
    try:
        put_yaml_object(client, bucket, config_key, config_data)
        return {
            "status": "uploaded",
            "message": f"Config uploaded to r2://{bucket}/{config_key}",
            "path": config_key,
            "scrapers_count": len(config_data.get("scrapers", []))
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Failed to upload config: {e}",
            "path": config_key
        }


def get_config_for_workflow(
    r2_prefix: str,
    client = None,
    bucket: Optional[str] = None,
) -> Dict:
    """
    Get config - ALWAYS from local file, never from R2.
    This is the main function used by build_monitor_stats and build_monitor_report.
    """
    try:
        return load_local_config()
    except FileNotFoundError:
        log.error(f"Local config not found at {LOCAL_CONFIG_PATH}")
        return {}


# ── STATS FUNCTIONS ──────────────────────────────────────────────────────

def build_monitor_stats(
    r2_prefix: str,
    days_lookback: int = 30,
    update_existing: bool = True,
    client = None,
    bucket: Optional[str] = None,
) -> Dict:
    """Build monitor_stats for a site by scanning its R2 data."""
    if client is None or bucket is None:
        client, bucket = build_r2_client()
    
    if not bucket:
        raise ValueError("Bucket name is required")
    
    # Load site config if exists
    try:
        site = load_site_config_from_r2(client, bucket, r2_prefix)
        log.info(f"Loaded site config for {r2_prefix}")
    except FileNotFoundError:
        log.info(f"No site config found for {r2_prefix}, using minimal config")
        site = {
            "folder": r2_prefix,
            "r2_prefix": r2_prefix,
            "site_id": r2_prefix,
        }
    
    # Get monitor paths using monitor_r2
    keys = monitor_data_keys(site)
    stats_key = keys["stats"]
    
    log.info(f"Site: {site.get('site_id')} · data prefix: {r2_prefix}")
    log.info(f"Stats: r2://{bucket}/{stats_key}")
    
    # Load config from LOCAL file ONLY
    config = get_config_for_workflow(r2_prefix, client, bucket)
    if not config:
        log.warning("No local config found, using minimal config")
        config = {}
    
    # Load existing stats
    existing_stats = {}
    if update_existing:
        try:
            existing_stats = fetch_yaml_object(client, bucket, stats_key)
            log.info(f"Loaded existing stats from {stats_key}")
        except Exception:
            log.info("No existing stats found, starting fresh")
    
    scrapers = config.get("scrapers", [])
    if not scrapers:
        log.warning("No scrapers found in config")
        return {}
    
    end_date = datetime.utcnow()
    dates_to_check = [
        (end_date - timedelta(days=i)).strftime("%Y-%m-%d")
        for i in range(days_lookback)
    ]
    print(dates_to_check)
    
    log.info(f"Processing {len(scrapers)} scrapers over {days_lookback} days...")
    
    stats = dict(existing_stats)
    
    for scraper_config in scrapers:
        scraper_name = scraper_config.get("name")
        if not scraper_name:
            continue
        
        log.info(f"Processing {scraper_name}...")
        
        scraper_stats = get_stats_for_scraper(
            client,
            bucket,
            scraper_name,
            scraper_config,
            dates_to_check,
            existing_stats
        )
        
        if scraper_stats:
            if scraper_name in stats:
                stats[scraper_name] = merge_stats(stats[scraper_name], scraper_stats)
            else:
                stats[scraper_name] = scraper_stats
    
    # Save stats using monitor_r2
    log.info(f"Saving stats to {stats_key}...")
    put_yaml_object(client, bucket, stats_key, stats)
    
    return stats


# ── REPORT FUNCTIONS ──────────────────────────────────────────────────────

def build_monitor_report(
    r2_prefix: str,
    target_date: Optional[str] = None,
    save_to_r2: bool = True,
    client = None,
    bucket: Optional[str] = None,
) -> Dict[str, Any]:
    """Build a monitor report."""
    if client is None or bucket is None:
        client, bucket = build_r2_client()
    
    if not bucket:
        raise ValueError("Bucket name is required")
    
    if target_date:
        dt = datetime.strptime(target_date, "%Y-%m-%d")
    else:
        dt = datetime.utcnow() - timedelta(days=1)
    
    try:
        site = load_site_config_from_r2(client, bucket, r2_prefix)
    except FileNotFoundError:
        site = {
            "folder": r2_prefix,
            "r2_prefix": r2_prefix,
            "site_id": r2_prefix,
        }
    
    config = get_config_for_workflow(r2_prefix, client, bucket)
    if not config:
        config = {}
    
    scrapers = config.get("scrapers", [])
    
    report = {
        "scraped_at": datetime.utcnow().isoformat(),
        "saved_to_s3_date": dt.strftime("%Y-%m-%d"),
        "site_id": site.get("site_id"),
        "site_name": site.get("display_name") or site.get("folder"),
        "r2_prefix": r2_prefix,
    }
    
    categories = get_categories_ads(
        client, bucket, scrapers, dt, download_files=True
    )
    
    categories = sorted(categories, key=lambda x: x.get("total_ads", 0), reverse=True)
    total_ads = sum(c.get("total_ads", 0) for c in categories)
    
    report.update({
        "total_categories": len(categories),
        "total_ads": total_ads,
        "categories": categories,
    })
    
    report["request_metrics"] = collect_request_metrics(
        client, bucket, scrapers, dt
    )
    
    try:
        total_files = count_site_r2_files(client, bucket, r2_prefix)
        report["total_r2_files"] = total_files
    except Exception:
        pass
    
    if save_to_r2:
        partition_date = dt.strftime("%Y-%m-%d")
        report_key = report_r2_key(site, partition_date)
        
        try:
            report_body = json.dumps(report, ensure_ascii=False, indent=2).encode('utf-8')
            client.put_object(
                Bucket=bucket,
                Key=report_key,
                Body=report_body,
                ContentType="application/json"
            )
        except Exception as e:
            log.error(f"Failed to save report to R2: {e}")
    
    return report


def build_monitor_report_with_date_range(
    r2_prefix: str,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    days_lookback: int = 30,
    client = None,
    bucket: Optional[str] = None,
) -> Dict[str, Any]:
    """Build a monitor report for a date range with aggregated stats."""
    if client is None or bucket is None:
        client, bucket = build_r2_client()
    
    if end_date:
        end_dt = datetime.strptime(end_date, "%Y-%m-%d")
    else:
        end_dt = datetime.utcnow() - timedelta(days=1)
    
    if start_date:
        start_dt = datetime.strptime(start_date, "%Y-%m-%d")
    else:
        start_dt = end_dt - timedelta(days=days_lookback)
    
    daily_reports = []
    current_dt = start_dt
    while current_dt <= end_dt:
        try:
            daily_report = build_monitor_report(
                r2_prefix=r2_prefix,
                target_date=current_dt.strftime("%Y-%m-%d"),
                save_to_r2=False,
                client=client,
                bucket=bucket
            )
            daily_reports.append(daily_report)
        except Exception as e:
            log.warning(f"Failed to build report for {current_dt.strftime('%Y-%m-%d')}: {e}")
        current_dt += timedelta(days=1)
    
    category_totals = defaultdict(int)
    all_categories = {}
    
    for report in daily_reports:
        for cat in report.get("categories", []):
            name = cat.get("name")
            slug = cat.get("slug", name.lower().replace(" ", "-"))
            category_totals[name] += cat.get("total_ads", 0)
            all_categories[name] = slug
    
    categories = []
    for name, slug in all_categories.items():
        categories.append({
            "name": name,
            "slug": slug,
            "total_ads": category_totals[name],
        })
    
    categories = sorted(categories, key=lambda x: x["total_ads"], reverse=True)
    
    return {
        "date_range": {
            "start": start_dt.strftime("%Y-%m-%d"),
            "end": end_dt.strftime("%Y-%m-%d"),
        },
        "site_id": r2_prefix,
        "site_name": r2_prefix,
        "r2_prefix": r2_prefix,
        "generated_at": datetime.utcnow().isoformat(),
        "total_categories": len(categories),
        "total_ads": sum(c["total_ads"] for c in categories),
        "categories": categories,
        "days_processed": len(daily_reports),
    }


# ── CLI ──────────────────────────────────────────────────────────────────

def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Monitor data management - config always from local file",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Upload local config to R2
  python build_monitor_data.py --prefix DOMAN --upload-config

  # Show local config
  python build_monitor_data.py --prefix DOMAN --show-config

  # Build stats (uses local config)
  python build_monitor_data.py --prefix DOMAN --build-stats --days 30

  # Build report (uses local config)
  python build_monitor_data.py --prefix DOMAN --report-date 2026-07-28

  # All in one
  python build_monitor_data.py --prefix DOMAN --build-stats --report-date 2026-07-28
        """
    )
    
    # Config options
    parser.add_argument("--prefix", required=True, help="R2 prefix (e.g., DKSA or 4sale-data)")
    parser.add_argument("--show-config", action="store_true", help="Show local config")
    parser.add_argument("--upload-config", action="store_true", help="Upload local config to R2")
    parser.add_argument("--force", action="store_true", help="Force upload (always true for local config)")
    
    # Stats options
    parser.add_argument("--build-stats", action="store_true", help="Build monitor_stats")
    parser.add_argument("--days", type=int, default=30, help="Days to look back for stats")
    parser.add_argument("--no-save", action="store_true", help="Don't save stats to R2")
    parser.add_argument("--output-stats", help="Save stats to local YAML file")
    
    # Report options
    parser.add_argument("--report-date", help="Date for report (YYYY-MM-DD), defaults to yesterday")
    parser.add_argument("--range-report", action="store_true", help="Build date range report")
    parser.add_argument("--start-date", help="Start date for range report (YYYY-MM-DD)")
    parser.add_argument("--end-date", help="End date for range report (YYYY-MM-DD)")
    parser.add_argument("--output-report", help="Save report to local JSON file")
    
    args = parser.parse_args()
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )
    
    client, bucket = build_r2_client()
    
    # ── Config ────────────────────────────────────────────────────────────
    if args.show_config:
        try:
            config = load_local_config()
            print("=" * 60)
            print(f"📄 LOCAL CONFIG: {LOCAL_CONFIG_PATH}")
            print(f"📊 Scrapers: {len(config.get('scrapers', []))}")
            print("=" * 60)
            print(yaml.dump(config, allow_unicode=True, default_flow_style=False))
        except FileNotFoundError as e:
            print(f"❌ {e}")
        return
    
    if args.upload_config:
        result = upload_config_to_r2(
            r2_prefix=args.prefix,
            force=True,
            client=client,
            bucket=bucket
        )
        print(result["message"])
        if result["status"] == "uploaded":
            print(f"   Scrapers: {result['scrapers_count']}")
        return
    
    # ── Stats ─────────────────────────────────────────────────────────────
    if args.build_stats:
        log.info(f"Building monitor_stats for {args.prefix}...")
        stats = build_monitor_stats(
            r2_prefix=args.prefix,
            days_lookback=args.days,
            update_existing=not args.no_save,
            client=client,
            bucket=bucket
        )
        
        if args.output_stats:
            with open(args.output_stats, "w", encoding="utf-8") as f:
                yaml.dump(stats, f, allow_unicode=True, default_flow_style=False, sort_keys=True)
            log.info(f"Stats saved to {args.output_stats}")
        
        if not args.report_date and not args.range_report:
            return
    
    # ── Report ────────────────────────────────────────────────────────────
    if args.report_date or args.range_report or args.start_date:
        if args.range_report or args.start_date:
            log.info(f"Building date range report for {args.prefix}...")
            report = build_monitor_report_with_date_range(
                r2_prefix=args.prefix,
                start_date=args.start_date,
                end_date=args.end_date,
                days_lookback=args.days,
                client=client,
                bucket=bucket
            )
        else:
            log.info(f"Building report for {args.report_date}...")
            report = build_monitor_report(
                r2_prefix=args.prefix,
                target_date=args.report_date,
                save_to_r2=not args.no_save,
                client=client,
                bucket=bucket
            )
        
        if args.output_report:
            with open(args.output_report, "w", encoding="utf-8") as f:
                json.dump(report, f, indent=2, ensure_ascii=False)
            log.info(f"Report saved to {args.output_report}")
        else:
            print(json.dumps(report, indent=2, ensure_ascii=False))
    
    if not any([args.show_config, args.upload_config, args.build_stats, args.report_date, args.range_report, args.start_date]):
        parser.print_help()


if __name__ == "__main__":
    main()