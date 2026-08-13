#!/usr/bin/env python3
"""
Finalize summaries: aggregate batch summaries (per category_path), add the
real workflow duration, and upload one summary.json per category to R2.
Works for Motors, Property, and Classifieds & Community workflows.
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
    Aggregate multiple batch summaries (same category_path) into one.
    request_metrics.requests_total / requests_failed are SUMMED across
    files too -- each batch already carries its own dub_writer-computed
    stats, and those must not be dropped when combining batches.
    """
    if not summary_files:
        return {}

    if len(summary_files) == 1:
        return load_summary(summary_files[0])

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
    requests_total_sum = 0
    requests_failed_sum = 0

    for filepath in summary_files:
        s = load_summary(filepath)

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

        rm = s.get("request_metrics", {}) or {}
        requests_total_sum += rm.get("requests_total", 0) or 0
        requests_failed_sum += rm.get("requests_failed", 0) or 0

    aggregated["total_listings"] = total_listings
    aggregated["total_subcategories"] = len(subcats)
    aggregated["subcategories"] = list(subcats.values())
    aggregated["request_metrics"] = {
        "requests_total": requests_total_sum,
        "requests_failed": requests_failed_sum,
    }

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


def finalize_summaries(summaries_dir: str, workflow_name: str = None, aggregate: bool = False):
    """
    For each category_path found under summaries_dir: aggregate its batch
    summaries (if more than one), add the real workflow duration, and
    upload ONE summary.json per category to R2 -- request_metrics and
    failed_items stay embedded inside that same file, never as separate
    uploads.
    """
    dt = datetime.now(timezone.utc)
    date_prefix = f"year={dt.year}/month={dt.strftime('%m')}/day={dt.strftime('%d')}"

    workflow_duration = os.getenv("WORKFLOW_DURATION")
    if not workflow_duration:
        print("\u26a0\ufe0f WORKFLOW_DURATION not set. Using fallback 0.")
        workflow_duration = "0"

    try:
        duration_sec = float(workflow_duration)
    except ValueError:
        duration_sec = 0.0

    print(f"\u2705 Workflow duration: {duration_sec}s")

    patterns = [
        os.path.join(summaries_dir, "*.json"),
        os.path.join(summaries_dir, "**", "summary.json"),
    ]

    summary_files = set()
    for pattern in patterns:
        for filepath in glob.glob(pattern, recursive=True):
            if not os.path.basename(filepath).startswith("request_stats_"):
                summary_files.add(filepath)

    summary_files = sorted(summary_files)

    if not summary_files:
        print(f"No summary files found in {summaries_dir}")
        return

    print(f"Found {len(summary_files)} summary file(s)")

    by_path = {}
    for filepath in summary_files:
        try:
            s = load_summary(filepath)
            cp = s.get("category_path")
            if not cp:
                basename = os.path.basename(filepath)
                if basename.startswith("summary_placeholder_"):
                    cat = basename.replace("summary_placeholder_", "").replace(".json", "")
                    cp = cat.replace("_", "/")
                else:
                    cp = "unknown"
            by_path.setdefault(cp, []).append(filepath)
        except Exception as e:
            print(f"  Could not read {filepath}: {e}")

    for category_path, files in by_path.items():
        print(f"\n  Processing: {category_path} ({len(files)} file(s))")

        summary = aggregate_summaries(files)
        if len(files) > 1:
            print(f"    Aggregated {len(files)} summaries")

        if "request_metrics" not in summary:
            summary["request_metrics"] = {}

        # Add the TRUE workflow duration (from the GH Actions run itself,
        # not any per-job/per-batch estimate) and recompute requests_per_min
        # from it -- this is the only thing finalize adds; requests_total /
        # requests_failed come straight from dub_writer.py's own numbers.
        summary["request_metrics"]["workflow_duration_sec"] = duration_sec

        total_requests = summary["request_metrics"].get("requests_total", 0)
        if duration_sec > 0:
            summary["request_metrics"]["requests_per_min"] = round(total_requests / (duration_sec / 60.0), 2)
        else:
            summary["request_metrics"]["requests_per_min"] = total_requests

        total_failed = summary["request_metrics"].get("requests_failed", len(summary.get("failed_items", [])))
        summary["request_metrics"]["requests_failed"] = total_failed
        if total_requests > 0:
            summary["request_metrics"]["error_rate_pct"] = round(total_failed / total_requests * 100, 2)
        else:
            summary["request_metrics"]["error_rate_pct"] = None

        if workflow_name:
            summary["workflow_name"] = workflow_name

        summary["failed_items_summary"] = format_failed_summary(summary.get("failed_items", []))

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
                print(f"    \u2705 Uploaded: {result}")
            else:
                print(f"    \u26a0\ufe0f Upload returned None for: {r2_key}")
        except Exception as e:
            print(f"    \u274c Upload failed: {e}")
            print(f"    Would upload to: {r2_key}")

    print(f"\n\U0001f389 Done! Processed {len(by_path)} category(ies).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Finalize summaries: aggregate per category_path, add real workflow duration, upload to R2"
    )
    parser.add_argument("--summaries-dir", default="summaries/", help="Directory containing summary JSON files")
    parser.add_argument("--workflow", default=None, help="Workflow name (e.g., 'Sale Property', 'Motors')")
    parser.add_argument("--aggregate", action="store_true", help="Aggregate multiple summaries per category_path (for Property)")
    args = parser.parse_args()

    finalize_summaries(args.summaries_dir, args.workflow, args.aggregate)