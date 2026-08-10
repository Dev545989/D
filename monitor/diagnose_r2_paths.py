"""
diagnose_r2_paths.py
=====================
Quick standalone check: what actually exists under DKSA/year=.../month=.../day=.../
in R2 right now, for a given date (or a few recent dates).

Run it directly in the GitHub Actions runner (or locally with the same env
vars) to compare the REAL category folder names against what
websites-config.yml expects.

Usage:
    python diagnose_r2_paths.py --prefix DKSA --date 2026-08-02
    python diagnose_r2_paths.py --prefix DKSA --days 5   # scans last 5 days
"""

import argparse
import os
from datetime import datetime, timedelta

import boto3


def build_client():
    access_key = os.environ["CF_R2_ACCESS_KEY_ID"]
    secret_key = os.environ["CF_R2_SECRET_ACCESS_KEY"]
    endpoint = os.environ["CF_R2_ENDPOINT_URL"].rstrip("/")
    bucket_name = os.environ["CF_R2_BUCKET_NAME"]

    if endpoint.endswith("/" + bucket_name):
        endpoint = endpoint[: -len("/" + bucket_name)]

    client = boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name="auto",
    )
    return client, bucket_name


def list_immediate_children(client, bucket, prefix):
    """List the folder names directly under `prefix` (one level deep)."""
    paginator = client.get_paginator("list_objects_v2")
    children = set()
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix, Delimiter="/"):
        for cp in page.get("CommonPrefixes", []):
            # cp["Prefix"] looks like "DKSA/year=2026/month=08/day=02/Vehicles/"
            child = cp["Prefix"][len(prefix):].rstrip("/")
            children.add(child)
    return sorted(children)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--prefix", default="DKSA")
    parser.add_argument("--date", help="Single date YYYY-MM-DD to check")
    parser.add_argument("--days", type=int, default=1, help="How many recent days to scan if --date not given")
    args = parser.parse_args()

    client, bucket = build_client()

    if args.date:
        dates = [datetime.strptime(args.date, "%Y-%m-%d")]
    else:
        today = datetime.utcnow()
        dates = [today - timedelta(days=i) for i in range(args.days)]

    for dt in dates:
        date_prefix = f"{args.prefix}/year={dt.year}/month={dt.month:02d}/day={dt.day:02d}/"
        print("=" * 70)
        print(f"📅 {dt.strftime('%Y-%m-%d')}  →  {date_prefix}")
        print("=" * 70)

        categories = list_immediate_children(client, bucket, date_prefix)
        if not categories:
            print("   ❌ NOTHING found under this date prefix at all.")
            continue

        print(f"   Found {len(categories)} folder(s) directly under this date:")
        for cat in categories:
            cat_prefix = f"{date_prefix}{cat}/"
            sub = list_immediate_children(client, bucket, cat_prefix)
            print(f"   📁 '{cat}'  → subfolders: {sub}")


if __name__ == "__main__":
    main()