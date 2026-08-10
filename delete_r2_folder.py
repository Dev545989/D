import os
import boto3
from botocore.client import Config
from dotenv import load_dotenv

load_dotenv()

CF_R2_ACCESS_KEY = os.getenv("CF_R2_ACCESS_KEY_ID")
CF_R2_SECRET_KEY = os.getenv("CF_R2_SECRET_ACCESS_KEY")
CF_R2_ENDPOINT_URL = os.getenv("CF_R2_ENDPOINT_URL")
BUCKET_NAME = os.getenv("CF_R2_BUCKET_NAME")

FOLDER = "DUAE/"

client = boto3.client(
    "s3",
    endpoint_url=CF_R2_ENDPOINT_URL,
    aws_access_key_id=CF_R2_ACCESS_KEY,
    aws_secret_access_key=CF_R2_SECRET_KEY,
    config=Config(signature_version="s3v4"),
    region_name="auto",
)

# ==========================================
# 1. Get EVERYTHING under DUAE/
# ==========================================

keys = []
paginator = client.get_paginator("list_objects_v2")

for page in paginator.paginate(
    Bucket=BUCKET_NAME,
    Prefix=FOLDER
):
    for obj in page.get("Contents", []):
        keys.append(obj["Key"])

print(f"Target: {FOLDER}")
print(f"Found {len(keys)} objects")

# ==========================================
# 2. Delete everything in batches of 1000
# ==========================================

if keys:

    deleted_count = 0
    error_count = 0

    for i in range(0, len(keys), 1000):

        batch = keys[i:i + 1000]

        response = client.delete_objects(
            Bucket=BUCKET_NAME,
            Delete={
                "Objects": [{"Key": key} for key in batch],
                "Quiet": False
            }
        )

        deleted = response.get("Deleted", [])
        errors = response.get("Errors", [])

        deleted_count += len(deleted)
        error_count += len(errors)

        print(
            f"Deleted {deleted_count} / {len(keys)} "
            f"| Errors: {error_count}"
        )

        if errors:
            for error in errors:
                print(
                    f"ERROR: {error.get('Key')} "
                    f"| {error.get('Code')} "
                    f"| {error.get('Message')}"
                )

    print("\n========== RESULT ==========")
    print(f"Deleted: {deleted_count}")
    print(f"Errors: {error_count}")

else:
    print("No files found under DUAE/")

# ==========================================
# 3. Verify that DUAE/ is empty
# ==========================================

remaining = []

for page in paginator.paginate(
    Bucket=BUCKET_NAME,
    Prefix=FOLDER
):
    for obj in page.get("Contents", []):
        remaining.append(obj["Key"])

print(f"Remaining: {len(remaining)}")

if not remaining:
    print("✅ DUAE/ is completely empty.")
else:
    print("❌ Some objects are still remaining.")

    for key in remaining[:20]:
        print(key)

print("Done")