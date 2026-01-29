import os
import uuid
from fastapi import APIRouter, UploadFile, File, HTTPException
from minio import Minio
from minio.error import S3Error
from .schemas import FileResponse, DeleteResponse, FileMetadata, ListFilesResponse

router = APIRouter()

MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "localhost:9000")
MINIO_ACCESS_KEY = os.getenv("MINIO_ACCESS_KEY", "minioadmin")
MINIO_SECRET_KEY = os.getenv("MINIO_SECRET_KEY", "minioadmin")
MINIO_BUCKET_NAME = os.getenv("MINIO_BUCKET_NAME", "uploads")

minio_client = Minio(MINIO_ENDPOINT, access_key=MINIO_ACCESS_KEY, secret_key=MINIO_SECRET_KEY, secure=False)

try:
    if not minio_client.bucket_exists(MINIO_BUCKET_NAME):
        minio_client.make_bucket(MINIO_BUCKET_NAME)
except S3Error as e:
    print(f"error ensuring bucket exists: {e}")


@router.post("/upload", response_model=FileResponse, summary="Upload a PowerPoint file")
async def upload_file(file: UploadFile = File(..., description="The .pptx file to upload")):
    """
    Upload a PowerPoint presentation file (.pptx) to the storage.
    Returns the generated file ID which can be used for further operations.
    """
    try:
        file_id = str(uuid.uuid4())
        metadata = {"filename": file.filename}
        minio_client.put_object(MINIO_BUCKET_NAME, file_id, file.file, length=-1, part_size=10 * 1024 * 1024, metadata=metadata)
        return FileResponse(id=file_id, filename=file.filename, message="file uploaded successfully")
    except Exception as e:
        return FileResponse(id="", filename="", message=f"error: {str(e)}")


@router.get("/list", response_model=ListFilesResponse, summary="List all uploaded files")
async def list_files():
    """
    Get a list of all uploaded file IDs in the storage bucket.
    """
    try:
        objects = minio_client.list_objects(MINIO_BUCKET_NAME)
        return ListFilesResponse(files=[obj.object_name for obj in objects])
    except Exception as e:
        return ListFilesResponse(files=[])


@router.get("/check-metadata/{file_id}", response_model=FileMetadata, summary="Get file metadata")
async def check_metadata(file_id: str):
    """
    Retrieve metadata for a specific file including filename, size, and content type.
    """
    try:
        stat = minio_client.stat_object(MINIO_BUCKET_NAME, file_id)
        filename = stat.metadata.get("x-amz-meta-filename", file_id)
        return FileMetadata(
            id=file_id, filename=filename, size=stat.size, content_type=stat.content_type, last_modified=stat.last_modified
        )
    except Exception as e:
        raise HTTPException(status_code=404, detail="file not found")


@router.delete("/delete/{file_id}", response_model=DeleteResponse, summary="Delete a file")
async def delete_file(file_id: str):
    """
    Permanently delete a file from the storage bucket.
    """
    try:
        minio_client.remove_object(MINIO_BUCKET_NAME, file_id)
        return DeleteResponse(id=file_id, message="file deleted successfully")
    except Exception as e:
        return DeleteResponse(id=file_id, message=f"error: {str(e)}")
