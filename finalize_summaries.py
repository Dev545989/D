#!/usr/bin/env python3
"""
Finalize summaries: aggregate batch summaries, add workflow duration, and upload to R2.
Works for both Property and Motors workflows.
"""

import argparse
import json
import os
import glob
import io
from datetime import datetime, timezone

from r2_uploader import upload_buffer


def load_summary(filepath: str) -> dict:
    """Load a summary JSON file."""
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


def aggregate_summaries(summary_files: list) -> dict:
    """
    Aggregate multiple batch summaries into one final summary.
    Used when a category has multiple batch summaries (e.g., Property with many batches).
    """
    if not summary_files:
        return {}

    # If only one file, return it directly
    if len(summary_files) == 1:
        return load_summary(summary_files[0])

    # Aggregate multiple summaries
    aggregated = {
        "scraped_at": None,
        "data_scraped_date": None,
        "saved_to_R2_date": None,
        "category": {},
        "category_path": None,
        "workflow_name": None,
        "total_subcategories": 0,
        "total_listings": 0,
        "subcategories": [],
        "request_metrics": {},
        "failed_items": [],
        "failed_items_summary": None,
    }

    subcats = {}
    total_listings = 0
    all_failed_items = []

    for filepath in summary_files:
        s = load_summary(filepath)

        # Use first summary for metadata
        if aggregated["scraped_at"] is None:
            aggregated["scraped_at"] = s.get("scraped_at")
            aggregated["data_scraped_date"] = s.get("data_scraped_date")
            aggregated["saved_to_R2_date"] = s.get("saved_to_R2_date")
            aggregated["category"] = s.get("category", {})
            aggregated["category_path"] = s.get("category_path")
            aggregated["workflow_name"] = s.get("workflow_name")

        total_listings += s.get("total_listings", 0)

        for sc in s.get("subcategories", []):
            key = sc.get("slug", sc.get("name_en", "unknown"))
            if key in subcats:
                subcats[key]["listings_count"] += sc.get("listings_count", 0)
            else:
                subcats[key] = dict(sc)

        all_failed_items.extend(s.get("failed_items", []))

    aggregated["total_listings"] = total_listings
    aggregated["total_subcategories"] = len(subcats)
    aggregated["subcategories"] = list(subcats.values())

    # Deduplicate failed items
    seen = set()
    unique_failed = []
    for item in all_failed_items:
        key = item.get("name", "")
        if key and key not in seen:
            seen.add(key)
            unique_failed.append(item)
    aggregated["failed_items"] = unique_failed

    return aggregated


def format_failed_summary(failed_items: list, max_len: int = 400) -> str | None:
    """Format failed items into a short summary string."""
    if not failed_items:
        return None
    parts = []
    for item in failed_items[:12]:
        name = item.get("name", "?")
        count = item.get("errors", 0)
        detail = item.get("detail", "")
        bit = f"{name}: {count} error(s)"
        if detail:
            bit += f" ({detail})"
        parts.append(bit)
    text = "; ".join(parts)
    if len(failed_items) > 12:
        text += f"; +{len(failed_items) - 12} more"
    return text[:max_len]


def aggregate_request_stats(summaries_dir: str) -> dict:
    """Read and aggregate all request_stats_*.json files in the directory."""
    stats_files = glob.glob(os.path.join(summaries_dir, "request_stats_*.json"))

    total_requests = 0
    total_duration_min = 0.0

    for sf in stats_files:
        try:
            with open(sf, 'r', encoding='utf-8') as f:
                st = json.load(f)
            total_requests += st.get('total_requests', 0)
            total_duration_min += st.get('total_duration_min', 0) or 0
        except Exception:
            pass

    result = {
        "requests_total": total_requests,
        "duration_sec": round(total_duration_min * 60, 2),
    }

    if total_duration_min > 0:
        result["requests_per_min"] = round(total_requests / total_duration_min, 2)
    else:
        result["requests_per_min"] = total_requests

    return result


def finalize_summaries(summaries_dir: str, workflow_name: str = None, aggregate: bool = False):
    """
    Finalize summaries: add workflow duration and upload to R2.

    Args:
        summaries_dir: Directory containing summary JSON files
        workflow_name: Name of the workflow (e.g., "Sale Property", "Motors")
        aggregate: If True, aggregate multiple summaries per category_path
    """
    dt = datetime.now(timezone.utc)
    date_prefix = f"year={dt.year}/month={dt.strftime('%m')}/day={dt.strftime('%d')}"

    # Read workflow duration from environment
    workflow_duration = os.getenv("WORKFLOW_DURATION")
    if not workflow_duration:
        print("⚠️ WORKFLOW_DURATION not set. Using fallback 0.")
        workflow_duration = "0"

    try:
        duration_sec = float(workflow_duration)
    except ValueError:
        duration_sec = 0.0

    print(f"✅ Workflow duration: {duration_sec}s")

    # Find all summary files (both summary.json and summary_placeholder_*.json)
    patterns = [
        os.path.join(summaries_dir, "*.json"),
        os.path.join(summaries_dir, "**", "summary.json"),
    ]

    summary_files = set()
    for pattern in patterns:
        for filepath in glob.glob(pattern, recursive=True):
            # Skip request stats files
            if not os.path.basename(filepath).startswith("request_stats_"):
                summary_files.add(filepath)

    summary_files = sorted(summary_files)

    if not summary_files:
        print(f"No summary files found in {summaries_dir}")
        return

    print(f"Found {len(summary_files)} summary file(s)")

    # Group by category_path
    by_path = {}
    for filepath in summary_files:
        try:
            s = load_summary(filepath)
            cp = s.get("category_path")
            if not cp:
                # Fallback: try to infer from filename or category
                basename = os.path.basename(filepath)
                if basename.startswith("summary_placeholder_"):
                    cat = basename.replace("summary_placeholder_", "").replace(".json", "")
                    cp = cat.replace("_", "/")
                else:
                    cp = "unknown"
            by_path.setdefault(cp, []).append(filepath)
        except Exception as e:
            print(f"  Could not read {filepath}: {e}")

    # Aggregate request stats from all stats files
    request_stats = aggregate_request_stats(summaries_dir)

    for category_path, files in by_path.items():
        print(f"\n  Processing: {category_path} ({len(files)} file(s))")

        # Aggregate summaries if needed
        if aggregate and len(files) > 1:
            summary = aggregate_summaries(files)
            print(f"    Aggregated {len(files)} summaries")
        else:
            summary = aggregate_summaries(files)

        # Ensure request_metrics exists
        if "request_metrics" not in summary:
            summary["request_metrics"] = {}

        # Merge aggregated request stats
        summary["request_metrics"].update(request_stats)

        # Add workflow duration (this is the KEY step - added at the END)
        summary["request_metrics"]["workflow_duration_sec"] = duration_sec

        # Update workflow name
        if workflow_name:
            summary["workflow_name"] = workflow_name

        # Format failed items summary
        summary["failed_items_summary"] = format_failed_summary(summary.get("failed_items", []))

        # Calculate error_rate_pct if possible
        total_requests = summary["request_metrics"].get("requests_total", 0)
        total_failed = len(summary.get("failed_items", []))
        if total_requests > 0:
            summary["request_metrics"]["requests_failed"] = total_failed
            summary["request_metrics"]["error_rate_pct"] = round(total_failed / total_requests * 100, 2)
        else:
            summary["request_metrics"]["requests_failed"] = total_failed
            summary["request_metrics"]["error_rate_pct"] = None

        # Upload to R2
        summary_bytes = json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8")
        r2_key = f"DUAE/{date_prefix}/{category_path}/summary/summary.json"

        try:
            result = upload_buffer(
                io.BytesIO(summary_bytes),
                filename="summary.json",
                folder_name="DUAE",
                category="",
                file_type="summary",
                content_type="application/json",
                dt=dt,
                category_path=category_path,
            )
            if result:
                print(f"    ✅ Uploaded: {result}")
            else:
                print(f"    ⚠️ Upload returned None for: {r2_key}")
        except Exception as e:
            print(f"    ❌ Upload failed: {e}")
            print(f"    Would upload to: {r2_key}")

    print(f"\n🎉 Done! Processed {len(by_path)} category(ies).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Finalize summaries: aggregate, add workflow duration, and upload to R2"
    )
    parser.add_argument(
        "--summaries-dir", 
        default="summaries/", 
        help="Directory containing summary JSON files"
    )
    parser.add_argument(
        "--workflow", 
        default=None, 
        help="Workflow name (e.g., 'Sale Property', 'Motors')"
    )
    parser.add_argument(
        "--aggregate", 
        action="store_true", 
        help="Aggregate multiple summaries per category_path (for Property)"
    )
    args = parser.parse_args()

    finalize_summaries(args.summaries_dir, args.workflow, args.aggregate)