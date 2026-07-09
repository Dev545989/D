import os
from datetime import datetime
import boto3
import io
from dotenv import load_dotenv
load_dotenv()

CF_R2_ACCESS_KEY = os.getenv('CF_R2_ACCESS_KEY_ID')
CF_R2_SECRET_KEY = os.getenv('CF_R2_SECRET_ACCESS_KEY')
CF_R2_ENDPOINT_URL = os.getenv('CF_R2_ENDPOINT_URL')
BUCKET_NAME = os.getenv('CF_R2_BUCKET_NAME', '')

CLEAN_ENDPOINT = ""
if CF_R2_ENDPOINT_URL:
    CLEAN_ENDPOINT = CF_R2_ENDPOINT_URL.rstrip("/").removesuffix("/" + BUCKET_NAME)


def get_r2_client():
    if CF_R2_ACCESS_KEY and CF_R2_SECRET_KEY and CLEAN_ENDPOINT:
        try:
            return boto3.client(
                's3',
                endpoint_url=CLEAN_ENDPOINT,
                aws_access_key_id=CF_R2_ACCESS_KEY,
                aws_secret_access_key=CF_R2_SECRET_KEY,
                region_name='auto'
            )
        except Exception as e:
            print(f"Failed to initialize R2 Client: {e}")
            return None
    print("Warning: R2 Environment variables are missing.")
    return None

R2_CLIENT_INSTANCE = get_r2_client()


def build_r2_key(folder_name: str, category: str, file_type: str, filename: str,
                  dt: datetime = None, city: str = None, category_display: str = None) -> str:
    if dt is None:
        dt = datetime.now()

    year  = f"year={dt.year}"
    month = f"month={dt.strftime('%m')}"
    day   = f"day={dt.strftime('%d')}"

    if city and category_display:
        if file_type:
            return f"{folder_name}/{year}/{month}/{day}/{city}/{category_display}/{file_type}/{filename}"
        else:
            return f"{folder_name}/{year}/{month}/{day}/{city}/{category_display}/{filename}"

    if file_type:
        return f"{folder_name}/{year}/{month}/{day}/{category}/{file_type}/{filename}"
    else:
        return f"{folder_name}/{year}/{month}/{day}/{category}/{filename}"


def upload_buffer(
    buffer: io.BytesIO,
    filename: str,
    folder_name: str = "qatarsale",
    category: str = "",
    file_type: str = "images",
    content_type: str = "image/webp",
    dt: datetime = None,
    city: str = None,
    category_display: str = None
) -> str | None:
    client = R2_CLIENT_INSTANCE if R2_CLIENT_INSTANCE else get_r2_client()
    if not client or not BUCKET_NAME:
        return None

    r2_key = build_r2_key(folder_name, category, file_type, filename, dt, city, category_display)

    try:
        buffer.seek(0)
        client.upload_fileobj(
            buffer, BUCKET_NAME, r2_key,
            ExtraArgs={"ContentType": content_type}
        )
        return r2_key
    except Exception as e:
        print(f"  [ERROR] R2 upload failed for {filename}: {e}")
        return None