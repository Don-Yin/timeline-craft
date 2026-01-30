"""upload service api endpoints"""
import os
import uuid

import httpx
from fastapi import APIRouter, BackgroundTasks, File, UploadFile
from minio import Minio

from .schemas import DeleteResponse, FileMetadata, FileResponse, ListFilesResponse

router = APIRouter()

MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "localhost:9000")
MINIO_ACCESS_KEY = os.getenv("MINIO_ACCESS_KEY", "minioadmin")
MINIO_SECRET_KEY = os.getenv("MINIO_SECRET_KEY", "minioadmin")
MINIO_BUCKET_NAME = os.getenv("MINIO_BUCKET_NAME", "uploads")
WORKER_SERVICE_URL = os.getenv("WORKER_SERVICE_URL", "http://worker-service:8002")

minio_client = Minio(MINIO_ENDPOINT, access_key=MINIO_ACCESS_KEY, secret_key=MINIO_SECRET_KEY, secure=False)

if not minio_client.bucket_exists(MINIO_BUCKET_NAME):
    minio_client.make_bucket(MINIO_BUCKET_NAME)


async def trigger_pdf_pregeneration(file_id: str):
    """background task to trigger pdf pre-generation on worker service"""
    async with httpx.AsyncClient(timeout=120.0) as client:
        response = await client.post(f"{WORKER_SERVICE_URL}/pre-generate-pdf/{file_id}")
        print(f"pdf pre-generation triggered for {file_id}: {response.status_code}")


@router.post("/upload", response_model=FileResponse, summary="Upload a PowerPoint file")
async def upload_file(background_tasks: BackgroundTasks, file: UploadFile = File(..., description="The .pptx file to upload")):
    """upload a powerpoint presentation file to storage"""
    file_id = str(uuid.uuid4())
    metadata = {"filename": file.filename}
    minio_client.put_object(MINIO_BUCKET_NAME, file_id, file.file, length=-1, part_size=10 * 1024 * 1024, metadata=metadata)
    background_tasks.add_task(trigger_pdf_pregeneration, file_id)
    return FileResponse(id=file_id, filename=file.filename, message="file uploaded successfully")


@router.get("/list", response_model=ListFilesResponse, summary="List all uploaded files")
async def list_files():
    """get a list of all uploaded file ids in the storage bucket"""
    objects = minio_client.list_objects(MINIO_BUCKET_NAME)
    return ListFilesResponse(files=[obj.object_name for obj in objects])


@router.get("/check-metadata/{file_id}", response_model=FileMetadata, summary="Get file metadata")
async def check_metadata(file_id: str):
    """retrieve metadata for a specific file including filename, size, and content type"""
    stat = minio_client.stat_object(MINIO_BUCKET_NAME, file_id)
    filename = stat.metadata.get("x-amz-meta-filename", file_id)
    return FileMetadata(id=file_id, filename=filename, size=stat.size, content_type=stat.content_type, last_modified=stat.last_modified)


@router.delete("/delete/{file_id}", response_model=DeleteResponse, summary="Delete a file")
async def delete_file(file_id: str):
    """permanently delete a file from the storage bucket"""
    minio_client.remove_object(MINIO_BUCKET_NAME, file_id)
    return DeleteResponse(id=file_id, message="file deleted successfully")
