import os
import boto3
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

import boto3
import yaml

MONITOR_SITES_ROOT = os.environ.get("MONITOR_SITES_PREFIX", "monitor-sites").strip("/")

log = logging.getLogger("monitor")


def build_r2_client() -> Tuple:
    """Return a boto3 S3 client pointed at Cloudflare R2."""
    access_key  = os.environ["CF_R2_ACCESS_KEY_ID"]
    secret_key  = os.environ["CF_R2_SECRET_ACCESS_KEY"]
    endpoint    = os.environ["CF_R2_ENDPOINT_URL"].rstrip("/")
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


def monitor_data_keys(site: Dict) -> Dict[str, str]:
    """Paths under the site's data prefix (excel/csv storage area)."""
    base = f"{site.get('r2_prefix', '').strip('/')}/monitor"
    return {
        "base":   base,
        "config": f"{base}/websites-config.yml",
        "stats":  f"{base}/monitor_stats.yml",
    }

def report_r2_key(site: Dict, partition_date: str) -> str:
    """R2 key for daily report — *partition_date* matches scraper save folders (listing + 1 day)."""
    return f"{monitor_data_keys(site)['base']}/{partition_date}/report.json"


def partition_date_for_listing(listing_dt):
    """
    Legacy function - kept for compatibility.
    """
    return listing_dt + timedelta(days=1)

def site_config_r2_key(folder: str, root: str = MONITOR_SITES_ROOT) -> str:
    return f"{root}/{folder.strip('/')}/site.yml"

def resolve_site_folder(explicit: Optional[str] = None) -> str:
    """Folder slug under monitor-sites/ (from env or CLI --site-slug)."""
    if explicit:
        return explicit.strip("/")
    env_slug = os.environ.get("MONITOR_SITE_SLUG", "").strip()
    if env_slug:
        return env_slug
    raise EnvironmentError(
        "MONITOR_SITE_SLUG is required (e.g. 4sale, boshamlan, motorgy, bleems, kcsb, sheeel). "
        "Set it in the repo's monitor.yml workflow env block."
    )

def fetch_yaml_object(client, bucket: str, key: str) -> Dict:
    resp = client.get_object(Bucket=bucket, Key=key)
    return yaml.safe_load(resp["Body"].read().decode("utf-8")) or {}

def put_yaml_object(client, bucket: str, key: str, data: Dict, header: str = "") -> None:
    body = (header + yaml.dump(data, allow_unicode=True, default_flow_style=False, sort_keys=False)).encode("utf-8")
    client.put_object(Bucket=bucket, Key=key, Body=body, ContentType="text/yaml")
    log.info(f"Uploaded → r2://{bucket}/{key}")

def load_site_config_from_r2(
    client,
    bucket: str,
    folder: Optional[str] = None,
    root: str = MONITOR_SITES_ROOT,
) -> Dict:
    slug = resolve_site_folder(folder)
    key  = site_config_r2_key(slug, root)
    try:
        site = fetch_yaml_object(client, bucket, key)
        site["folder"] = slug
        log.info(f"Loaded site config from r2://{bucket}/{key}")
        return site
    except client.exceptions.NoSuchKey:
        raise FileNotFoundError(
            f"Site config not found at r2://{bucket}/{key}. "
            f"Create monitor-sites/{slug}/site.yml in R2 (Cloudflare dashboard or aws s3 cp)."
        ) from None