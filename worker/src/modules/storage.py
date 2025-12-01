import io
import logging
import os
from typing import Optional
from minio import Minio
from minio.error import S3Error
from fastapi import HTTPException

logger = logging.getLogger(__name__)

MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "localhost:9000")
MINIO_ACCESS_KEY = os.getenv("MINIO_ACCESS_KEY", "minioadmin")
MINIO_SECRET_KEY = os.getenv("MINIO_SECRET_KEY", "minioadmin")
MINIO_BUCKET_UPLOADS = os.getenv("MINIO_BUCKET_UPLOADS", "uploads")
MINIO_BUCKET_PROCESSED = os.getenv("MINIO_BUCKET_PROCESSED", "processed")

minio_client = Minio(MINIO_ENDPOINT, access_key=MINIO_ACCESS_KEY, secret_key=MINIO_SECRET_KEY, secure=False)


def ensure_bucket(bucket_name: str):
    try:
        if not minio_client.bucket_exists(bucket_name):
            minio_client.make_bucket(bucket_name)
    except S3Error as exc:
        logger.error(f"error ensuring bucket {bucket_name}: {exc}")


def get_file_from_minio(bucket: str, file_id: str) -> io.BytesIO:
    try:
        response = minio_client.get_object(bucket, file_id)
        file_data = io.BytesIO(response.read())
        response.close()
        response.release_conn()
        return file_data
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"File not found in {bucket}: {str(e)}")


def upload_file_to_minio(bucket: str, file_id: str, data: io.BytesIO, content_type: str):
    try:
        minio_client.put_object(bucket, file_id, data, length=data.getbuffer().nbytes, content_type=content_type)
    except Exception as e:
        logger.error(f"Error uploading file to {bucket}: {e}")
        raise HTTPException(status_code=500, detail=f"Error uploading file: {str(e)}")


def list_files_in_bucket(bucket: str) -> list[str]:
    try:
        objects = minio_client.list_objects(bucket)
        return [obj.object_name for obj in objects]
    except Exception as e:
        logger.error(f"Error listing files in {bucket}: {e}")
        return []


# Ensure buckets exist on module load
ensure_bucket(MINIO_BUCKET_PROCESSED)
# MINIO_BUCKET_PREVIEWS logic was removed in previous step in favor of memory cache,
# but we keep PROCESSED bucket ensure here.
