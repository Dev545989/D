"""
ONE-OFF migration: consolidate every historical day-scoped
profiles-data.xlsx (rent + sale) into a single de-duplicated snapshot.

Background: profiles_key used to be recomputed and re-read fresh every day
under that day's own folder, so the phone-number cache was silently reset to
empty every single day instead of accumulating. This script walks every day
that was ever uploaded, up to and including END_DATE, concatenates all the
profiles it finds, drops duplicate (profile_type, slug) pairs (keeping the
most recently-seen phone), and writes ONE consolidated snapshot back to
END_DATE's own day-scoped folder:

  DUAE/year=2026/month=09/day=21/property/property-for-rent/profiles-data/profiles-data.xlsx
  DUAE/year=2026/month=09/day=21/property/property-for-sale/profiles-data/profiles-data.xlsx

From day=22 onward, prepare()'s find_latest_profiles_key() walks backward
day by day looking for the closest existing snapshot -- so it will land on
this consolidated day=21 file automatically and keep building forward from
it. No other write is needed.

Run once, manually, via the migrate_profiles_cache.yml workflow_dispatch
workflow. Safe to re-run (idempotent): re-running just recomputes the same
de-duplicated union from the source day-scoped files, which this script
never deletes.
"""

from __future__ import annotations

import re
from datetime import date

import pandas as pd

from r2_property_enrichment import (
    DUAE_PREFIX,
    LISTING_TYPE_PATH_FRAGMENTS,
    PHONE_COLUMN,
    PROPERTY_ROOT,
    R2_BUCKET,
    build_excel_bytes,
    download_bytes,
    excel_sheets,
    is_empty,
    r2_client,
    upload_bytes,
)

# Inclusive cutoff -- matches the last day that was written the old
# (day-scoped, self-resetting) way. Change this if you need to re-run the
# migration later against a different cutoff.
END_DATE = date(2026, 9, 21)

KEY_PATTERN = re.compile(
    r"^"
    + re.escape(DUAE_PREFIX)
    + r"/year=(\d{4})/month=(\d{2})/day=(\d{2})/"
    + re.escape(PROPERTY_ROOT)
    + r"/(?P<fragment>property-for-rent|property-for-sale)/profiles-data/profiles-data\.xlsx$"
)

FRAGMENT_TO_LISTING_TYPE = {v: k for k, v in LISTING_TYPE_PATH_FRAGMENTS.items()}


def list_all_keys(client) -> list[str]:
    keys = []
    paginator = client.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=R2_BUCKET, Prefix=f"{DUAE_PREFIX}/"):
        for obj in page.get("Contents", []):
            keys.append(obj["Key"])
    return keys


def find_profile_files(client) -> dict[str, list[tuple[date, str]]]:
    """Returns {listing_type: [(day, key), ...]} sorted oldest -> newest,
    restricted to files at or before END_DATE."""
    found: dict[str, list[tuple[date, str]]] = {t: [] for t in LISTING_TYPE_PATH_FRAGMENTS}

    all_keys = list_all_keys(client)
    print(f"[SCAN] {len(all_keys)} total object(s) under {DUAE_PREFIX}/")

    for key in all_keys:
        m = KEY_PATTERN.match(key)
        if not m:
            continue
        year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            file_date = date(year, month, day)
        except ValueError:
            print(f"[SCAN][WARN] Skipping key with invalid date: {key}")
            continue
        if file_date > END_DATE:
            continue
        listing_type = FRAGMENT_TO_LISTING_TYPE[m.group("fragment")]
        found[listing_type].append((file_date, key))

    for listing_type in found:
        found[listing_type].sort(key=lambda pair: pair[0])

    return found


def load_profiles(client, key: str) -> pd.DataFrame:
    try:
        data = download_bytes(client, key)
    except Exception as exc:
        print(f"  [WARN] Could not read {key}: {exc}")
        return pd.DataFrame()

    sheets = excel_sheets(data)
    frames = [df for df in sheets.values() if "profile_type" in df.columns and "slug" in df.columns]
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def migrate_listing_type(client, listing_type: str, dated_keys: list[tuple[date, str]]) -> None:
    fragment = LISTING_TYPE_PATH_FRAGMENTS[listing_type]

    if not dated_keys:
        print(f"[{listing_type}] No historical profiles-data.xlsx files found -- nothing to do.")
        return

    print(
        f"[{listing_type}] Found {len(dated_keys)} day(s) of cache, "
        f"from {dated_keys[0][0].isoformat()} to {dated_keys[-1][0].isoformat()}."
    )

    all_frames = []
    for file_date, key in dated_keys:
        df = load_profiles(client, key)
        print(f"  {file_date.isoformat()}: {key} -> {len(df)} row(s)")
        if not df.empty:
            all_frames.append(df)

    if not all_frames:
        print(f"[{listing_type}] Nothing readable -- nothing to do.")
        return

    combined = pd.concat(all_frames, ignore_index=True)

    if PHONE_COLUMN in combined.columns:
        combined = combined[~combined[PHONE_COLUMN].apply(is_empty)]

    before = len(combined)
    combined = combined.drop_duplicates(subset=["profile_type", "slug"], keep="last")
    print(f"[{listing_type}] Combined {before} row(s) -> {len(combined)} unique profile(s) after de-dup.")

    keep_cols = [c for c in ["profile_type", "slug", PHONE_COLUMN, "updated_at"] if c in combined.columns]
    combined = combined[keep_cols]

    payload = build_excel_bytes({"profiles": combined})
    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    end_prefix = (
        f"{DUAE_PREFIX}/year={END_DATE.year}/month={END_DATE.month:02d}/day={END_DATE.day:02d}/{PROPERTY_ROOT}/"
    )
    archive_key = f"{end_prefix}{fragment}/profiles-data/profiles-data.xlsx"
    upload_bytes(client, archive_key, payload, content_type)
    print(f"[{listing_type}] Saved consolidated snapshot -> {archive_key}")


def main():
    client = r2_client()
    found = find_profile_files(client)
    for listing_type in LISTING_TYPE_PATH_FRAGMENTS:
        migrate_listing_type(client, listing_type, found[listing_type])


if __name__ == "__main__":
    main()