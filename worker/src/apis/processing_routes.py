"""presentation processing and download endpoints"""
import io
import json
import logging
import time
import uuid
from threading import Lock

from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import StreamingResponse

from .schemas import ProcessRequest, ProcessResponse, ListProcessedResponse
from .thumbnail_routes import pdf_cache
from .preview_routes import get_preview_cache
from ..modules.storage import get_file_from_minio, upload_file_to_minio, list_files_in_bucket, MINIO_BUCKET_UPLOADS, MINIO_BUCKET_PROCESSED
from ..modules.pptx_processor import process_presentation
from ..modules.progressive_processor import ProgressiveProcessor

logger = logging.getLogger(__name__)
router = APIRouter(tags=["processing"])

processed_files_store: dict[str, tuple[float, str, io.BytesIO]] = {}
store_lock = Lock()
STORE_TTL = 300


def cleanup_expired_files():
    """remove expired files from temporary store"""
    now = time.time()
    with store_lock:
        expired = [k for k, (ts, _, _) in processed_files_store.items() if now - ts > STORE_TTL]
        for k in expired:
            del processed_files_store[k]


def generate_progress_stream(file_id: str, request: ProcessRequest):
    """generator for sse progress events"""
    cleanup_expired_files()

    file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
    processor = ProgressiveProcessor(file_data, request)

    gen = processor.process_with_progress()
    result_buffer = None

    while True:
        try:
            progress_event = next(gen)
            event_data = json.dumps(progress_event)
            yield f"data: {event_data}\n\n"
        except StopIteration as e:
            result_buffer = e.value
            break

    job_id = str(uuid.uuid4())
    with store_lock:
        processed_files_store[job_id] = (time.time(), file_id, result_buffer)

    done_event = json.dumps({"stage": "ready", "progress": 100, "job_id": job_id, "file_id": file_id})
    yield f"data: {done_event}\n\n"


@router.post("/process-file/{file_id}", response_model=ProcessResponse, summary="Process and store presentation")
async def process_file(file_id: str, request: ProcessRequest):
    """process a presentation with timeline sidebar and store the result"""
    file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
    processed_buffer = process_presentation(file_data, request)
    upload_file_to_minio(
        MINIO_BUCKET_PROCESSED,
        file_id,
        processed_buffer,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    return ProcessResponse(file_id=file_id, processed_file_id=file_id, message="file processed successfully")


@router.get("/list-processed", response_model=ListProcessedResponse, summary="List processed files")
async def list_processed():
    """get a list of all processed file ids"""
    files = list_files_in_bucket(MINIO_BUCKET_PROCESSED)
    return ListProcessedResponse(files=files)


@router.post("/process-and-download/{file_id}", summary="Process and download directly")
async def process_and_download(file_id: str, request: ProcessRequest):
    """process a presentation and return it directly as a downloadable file"""
    logger.info(f"processing file {file_id} for download")
    file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
    processed_buffer = process_presentation(file_data, request)

    return StreamingResponse(
        processed_buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename=timeline-{file_id}.pptx"},
    )


@router.post("/process-with-progress/{file_id}", summary="Process with real-time progress")
async def process_with_progress(file_id: str, request: ProcessRequest):
    """process a presentation with sse for real-time progress updates"""
    logger.info(f"starting progressive processing for {file_id}")

    return StreamingResponse(
        generate_progress_stream(file_id, request),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive", "X-Accel-Buffering": "no"},
    )


@router.get("/download-processed/{job_id}", summary="Download processed file by job ID")
async def download_processed(job_id: str, file_id: str = Query(default="")):
    """download a processed file using the job id from process-with-progress"""
    cleanup_expired_files()

    with store_lock:
        entry = processed_files_store.get(job_id)
        if not entry:
            raise HTTPException(status_code=404, detail="processed file not found or expired")

        _, stored_file_id, buffer = entry
        if file_id and stored_file_id != file_id:
            raise HTTPException(status_code=404, detail="processed file not found or expired")
        buffer.seek(0)
        content = buffer.read()
        del processed_files_store[job_id]

    pdf_cache.delete(stored_file_id)
    get_preview_cache().delete_prefix(f"{stored_file_id}:")

    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename=timeline-{job_id}.pptx"},
    )

