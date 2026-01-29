import logging
import os
import io
import json
import uuid
import time
from threading import Lock

from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse
from pptx import Presentation

from .schemas import ThumbnailResponse, ProcessRequest, ProcessResponse, ListProcessedResponse, SlideCountResponse
from ..modules.cache import PdfCache
from ..modules.storage import (
    get_file_from_minio,
    upload_file_to_minio,
    list_files_in_bucket,
    MINIO_BUCKET_UPLOADS,
    MINIO_BUCKET_PROCESSED,
)
from ..modules.pdf_renderer import convert_pptx_to_pdf_bytes, render_slide_thumbnail
from ..modules.pptx_processor import process_presentation
from ..modules.progressive_processor import ProgressiveProcessor

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

router = APIRouter()

PREVIEW_CACHE_TTL = int(os.getenv("PREVIEW_CACHE_TTL", "300"))
pdf_cache = PdfCache(ttl_seconds=PREVIEW_CACHE_TTL)

processed_files_store: dict[str, tuple[float, io.BytesIO]] = {}
store_lock = Lock()
STORE_TTL = 300


@router.get(
    "/get-thumbnail/{file_id}/{slide_index}",
    response_model=ThumbnailResponse,
    tags=["thumbnails"],
    summary="Get slide thumbnail"
)
async def get_thumbnail(file_id: str, slide_index: int):
    """
    Generate a PNG thumbnail for a specific slide in a presentation.
    Returns a base64-encoded image string.
    """
    logger.info(f"fetching thumbnail for file {file_id}, slide {slide_index}")

    try:
        file_buffer = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
        pptx_bytes = file_buffer.getvalue()
        prs = Presentation(io.BytesIO(pptx_bytes))

        logger.info(f"loaded presentation with {len(prs.slides)} slides")

        if slide_index < 0 or slide_index >= len(prs.slides):
            logger.warning(f"slide index {slide_index} out of range")
            raise HTTPException(status_code=404, detail="Slide index out of range")

        pdf_bytes = pdf_cache.get(file_id)
        if pdf_bytes is None:
            logger.info(f"no cached pdf for {file_id}, converting via libreoffice")
            pdf_bytes = convert_pptx_to_pdf_bytes(file_id, pptx_bytes)
            pdf_cache.set(file_id, pdf_bytes)

        img_str = render_slide_thumbnail(pdf_bytes, slide_index)
        return ThumbnailResponse(slide_index=slide_index, image_base64=img_str)
    except HTTPException as e:
        raise e
    except Exception as e:
        logger.error(f"unexpected error: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Error generating thumbnail: {str(e)}")


@router.get(
    "/get-slide-count/{file_id}",
    response_model=SlideCountResponse,
    tags=["thumbnails"],
    summary="Get slide count"
)
async def get_slide_count(file_id: str):
    """
    Get the total number of slides in a presentation.
    """
    try:
        file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
        prs = Presentation(file_data)
        return SlideCountResponse(file_id=file_id, slide_count=len(prs.slides))
    except HTTPException as e:
        raise e
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error getting slide count: {str(e)}")


@router.post(
    "/process-file/{file_id}",
    response_model=ProcessResponse,
    tags=["processing"],
    summary="Process and store presentation"
)
async def process_file(file_id: str, request: ProcessRequest):
    """
    Process a presentation with timeline sidebar and store the result.
    The processed file is saved to the processed bucket.
    """
    try:
        file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
        processed_buffer = process_presentation(file_data, request)
        upload_file_to_minio(
            MINIO_BUCKET_PROCESSED,
            file_id,
            processed_buffer,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        return ProcessResponse(file_id=file_id, processed_file_id=file_id, message="File processed successfully")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Processing error: {str(e)}")


@router.get(
    "/list-processed",
    response_model=ListProcessedResponse,
    tags=["processing"],
    summary="List processed files"
)
async def list_processed():
    """
    Get a list of all processed file IDs.
    """
    files = list_files_in_bucket(MINIO_BUCKET_PROCESSED)
    return ListProcessedResponse(files=files)


@router.post(
    "/process-and-download/{file_id}",
    tags=["processing"],
    summary="Process and download directly"
)
async def process_and_download(file_id: str, request: ProcessRequest):
    """
    Process a presentation and return it directly as a downloadable file.
    This is a synchronous operation - the response is the processed .pptx file.
    """
    logger.info(f"processing file {file_id} for download")
    file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
    processed_buffer = process_presentation(file_data, request)
    
    return StreamingResponse(
        processed_buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename=timeline-{file_id}.pptx"}
    )


def _cleanup_expired_files():
    """remove expired files from temporary store"""
    now = time.time()
    with store_lock:
        expired = [k for k, (ts, _) in processed_files_store.items() if now - ts > STORE_TTL]
        for k in expired:
            del processed_files_store[k]


def _generate_progress_stream(file_id: str, request: ProcessRequest):
    """generator for SSE progress events"""
    _cleanup_expired_files()
    
    file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
    processor = ProgressiveProcessor(file_data, request)
    
    result_buffer = None
    
    for progress_event in processor.process_with_progress():
        if isinstance(progress_event, io.BytesIO):
            result_buffer = progress_event
            break
        
        event_data = json.dumps(progress_event)
        yield f"data: {event_data}\n\n"
        
        if progress_event.get("stage") == "complete":
            break
    
    if result_buffer is None:
        result_buffer = io.BytesIO()
        processor.prs.save(result_buffer)
        result_buffer.seek(0)
    
    job_id = str(uuid.uuid4())
    with store_lock:
        processed_files_store[job_id] = (time.time(), result_buffer)
    
    done_event = json.dumps({"stage": "ready", "progress": 100, "job_id": job_id})
    yield f"data: {done_event}\n\n"


@router.post(
    "/process-with-progress/{file_id}",
    tags=["processing"],
    summary="Process with real-time progress"
)
async def process_with_progress(file_id: str, request: ProcessRequest):
    """
    Process a presentation with Server-Sent Events (SSE) for real-time progress updates.
    
    The response is a stream of JSON events:
    - `{"stage": "loading", "progress": 5, "message": "loading..."}`
    - `{"stage": "moving", "progress": 30, "message": "repositioning slide 5/20"}`
    - `{"stage": "ready", "progress": 100, "job_id": "abc-123"}`
    
    Use the returned `job_id` to download the processed file via `/download-processed/{job_id}`.
    """
    logger.info(f"starting progressive processing for {file_id}")
    
    return StreamingResponse(
        _generate_progress_stream(file_id, request),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        }
    )


@router.get(
    "/download-processed/{job_id}",
    tags=["processing"],
    summary="Download processed file by job ID"
)
async def download_processed(job_id: str):
    """
    Download a processed file using the job ID returned from `/process-with-progress`.
    The file is available for 5 minutes after processing completes.
    """
    _cleanup_expired_files()
    
    with store_lock:
        entry = processed_files_store.get(job_id)
        if not entry:
            raise HTTPException(status_code=404, detail="Processed file not found or expired")
        
        _, buffer = entry
        buffer.seek(0)
        content = buffer.read()
        del processed_files_store[job_id]
    
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename=timeline-{job_id}.pptx"}
    )
