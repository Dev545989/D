import logging
log = logging.getLogger("monitor")

def count_r2_objects(client, bucket: str, prefix: str) -> int:
    """
    Count all objects under *prefix* using paginated list_objects_v2.

    Skips zero-byte folder marker keys ending with '/'.
    """
    normalized = prefix.strip("/")
    list_prefix = f"{normalized}/" if normalized else ""

    count = 0
    paginator = client.get_paginator("list_objects_v2")
    try:
        for page in paginator.paginate(Bucket=bucket, Prefix=list_prefix):
            for obj in page.get("Contents", []):
                key = obj["Key"]
                if key.endswith("/"):
                    continue
                count += 1
    except Exception as exc:
        log.warning(f"R2 object count failed for prefix {list_prefix!r}: {exc}")
        return 0

    return count

def count_site_r2_files(client, bucket: str, r2_prefix: str) -> int:
    """Total objects under the site's data prefix (all scrapers + monitor artifacts)."""
    prefix = r2_prefix.strip("/")
    if not prefix:
        return 0
    total = count_r2_objects(client, bucket, prefix)
    log.info(f"Site R2 inventory ({prefix}): {total} object(s)")
    return total