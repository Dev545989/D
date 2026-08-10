"""
diagnose_excel_files.py
========================
Lists the REAL object keys (not just folder names) under one scraper's
excel/ folder for a given date, and re-runs the exact same code path
build_monitor_data.py uses (r2_base_prefix -> excel_prefixes_for_date ->
list_excel_files) so we can see exactly what it finds (or doesn't).

Usage:
    python diagnose_excel_files.py --prefix DKSA --category Vehicles --date 2026-08-02
"""

import argparse
import os
from datetime import datetime

import boto3

from inspect_r2_schema import r2_base_prefix, excel_prefixes_for_date, list_excel_files


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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--prefix", default="DKSA")
    parser.add_argument("--category", required=True, help="e.g. Vehicles")
    parser.add_argument("--date", required=True, help="YYYY-MM-DD")
    args = parser.parse_args()

    client, bucket = build_client()
    dt = datetime.strptime(args.date, "%Y-%m-%d")

    # 1) Raw listing — exactly what's really there, no filtering at all.
    raw_prefix = f"{args.prefix}/year={dt.year}/month={dt.month:02d}/day={dt.day:02d}/{args.category}/excel/"
    print(f"RAW listing under: {raw_prefix}")
    print("-" * 70)
    paginator = client.get_paginator("list_objects_v2")
    count = 0
    for page in paginator.paginate(Bucket=bucket, Prefix=raw_prefix):
        for obj in page.get("Contents", []):
            print(f"   {obj['Key']}   ({obj['Size']} bytes)")
            count += 1
    print(f"-> {count} raw object(s) found.\n")

    # 2) Re-run the EXACT code path build_monitor_data.py uses.
    r2_path_raw = "{r2_bucket}/" + f"{args.prefix}/{args.category}"
    base, category = r2_base_prefix(r2_path_raw)
    print(f"r2_base_prefix('{r2_path_raw}') -> base={base!r}, category={category!r}")

    prefixes = excel_prefixes_for_date(base, category, dt)
    print(f"excel_prefixes_for_date(...) -> {prefixes}")

    for p in prefixes:
        files = list_excel_files(client, bucket, p)
        print(f"list_excel_files(..., {p!r}) -> {len(files)} file(s)")
        for f in files[:5]:
            print(f"   {f}")


if __name__ == "__main__":
    main()