"""minio storage utilities for file operations"""
import io
import logging
import os

from minio import Minio
from minio.error import S3Error

logger = logging.getLogger(__name__)

MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "localhost:9000")
MINIO_ACCESS_KEY = os.getenv("MINIO_ACCESS_KEY", "minioadmin")
MINIO_SECRET_KEY = os.getenv("MINIO_SECRET_KEY", "minioadmin")
MINIO_BUCKET_UPLOADS = os.getenv("MINIO_BUCKET_UPLOADS", "uploads")
MINIO_BUCKET_PROCESSED = os.getenv("MINIO_BUCKET_PROCESSED", "processed")
MINIO_BUCKET_PDF_CACHE = os.getenv("MINIO_BUCKET_PDF_CACHE", "pdf-cache")

minio_client = Minio(MINIO_ENDPOINT, access_key=MINIO_ACCESS_KEY, secret_key=MINIO_SECRET_KEY, secure=False)


def ensure_bucket(bucket_name: str):
    """create bucket if it doesn't exist"""
    if not minio_client.bucket_exists(bucket_name):
        minio_client.make_bucket(bucket_name)


def get_file_from_minio(bucket: str, file_id: str) -> io.BytesIO:
    """retrieve file from minio bucket"""
    response = minio_client.get_object(bucket, file_id)
    file_data = io.BytesIO(response.read())
    response.close()
    response.release_conn()
    return file_data


def upload_file_to_minio(bucket: str, file_id: str, data: io.BytesIO, content_type: str):
    """upload file to minio bucket"""
    minio_client.put_object(bucket, file_id, data, length=data.getbuffer().nbytes, content_type=content_type)


def list_files_in_bucket(bucket: str) -> list[str]:
    """list all files in a bucket"""
    objects = minio_client.list_objects(bucket)
    return [obj.object_name for obj in objects]


def file_exists_in_minio(bucket: str, file_id: str) -> bool:
    """check if a file exists in minio bucket"""
    try:
        minio_client.stat_object(bucket, file_id)
        return True
    except S3Error:
        return False


def get_cached_pdf(file_id: str) -> bytes | None:
    """get cached pdf from minio, returns none if not found"""
    pdf_key = f"{file_id}.pdf"
    if not file_exists_in_minio(MINIO_BUCKET_PDF_CACHE, pdf_key):
        return None
    response = minio_client.get_object(MINIO_BUCKET_PDF_CACHE, pdf_key)
    data = response.read()
    response.close()
    response.release_conn()
    logger.info(f"loaded cached pdf for {file_id}")
    return data


def store_cached_pdf(file_id: str, pdf_bytes: bytes):
    """store pdf in minio cache for fast thumbnail generation"""
    pdf_key = f"{file_id}.pdf"
    minio_client.put_object(
        MINIO_BUCKET_PDF_CACHE,
        pdf_key,
        io.BytesIO(pdf_bytes),
        length=len(pdf_bytes),
        content_type="application/pdf"
    )
    logger.info(f"cached pdf for {file_id} ({len(pdf_bytes)} bytes)")


ensure_bucket(MINIO_BUCKET_PROCESSED)
ensure_bucket(MINIO_BUCKET_PDF_CACHE)
