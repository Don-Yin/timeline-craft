import logging
import os
import io

from fastapi import APIRouter, HTTPException
from pptx import Presentation

from .schemas import ThumbnailResponse, ProcessRequest, ProcessResponse, ListProcessedResponse, SlideCountResponse
from ..modules.cache import PdfCache
from ..modules.storage import get_file_from_minio, upload_file_to_minio, list_files_in_bucket, MINIO_BUCKET_UPLOADS, MINIO_BUCKET_PROCESSED
from ..modules.pdf_renderer import convert_pptx_to_pdf_bytes, render_slide_thumbnail
from ..modules.pptx_processor import process_presentation

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

router = APIRouter()

PREVIEW_CACHE_TTL = int(os.getenv("PREVIEW_CACHE_TTL", "300"))
pdf_cache = PdfCache(ttl_seconds=PREVIEW_CACHE_TTL)


@router.get("/get-thumbnail/{file_id}/{slide_index}", response_model=ThumbnailResponse)
async def get_thumbnail(file_id: str, slide_index: int):
    logger.info(f"Fetching thumbnail for file {file_id}, slide {slide_index}")

    try:
        file_buffer = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
        pptx_bytes = file_buffer.getvalue()
        prs = Presentation(io.BytesIO(pptx_bytes))

        logger.info(f"Loaded presentation. Slide count: {len(prs.slides)}")

        if slide_index < 0 or slide_index >= len(prs.slides):
            logger.warning(f"Slide index {slide_index} out of range. Max index: {len(prs.slides) - 1}")
            raise HTTPException(status_code=404, detail="Slide index out of range")

        pdf_bytes = pdf_cache.get(file_id)
        if pdf_bytes is None:
            logger.info(f"No cached PDF for {file_id}, converting via LibreOffice")
            pdf_bytes = convert_pptx_to_pdf_bytes(file_id, pptx_bytes)
            pdf_cache.set(file_id, pdf_bytes)

        img_str = render_slide_thumbnail(pdf_bytes, slide_index)

        return ThumbnailResponse(slide_index=slide_index, image_base64=img_str)
    except HTTPException as e:
        logger.warning(f"HTTP error: {e.detail}")
        raise e
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Error generating thumbnail: {str(e)}")


@router.get("/get-slide-count/{file_id}", response_model=SlideCountResponse)
async def get_slide_count(file_id: str):
    try:
        file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)
        prs = Presentation(file_data)
        return SlideCountResponse(file_id=file_id, slide_count=len(prs.slides))
    except HTTPException as e:
        raise e
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error getting slide count: {str(e)}")


@router.post("/process-file/{file_id}", response_model=ProcessResponse)
async def process_file(file_id: str, request: ProcessRequest):
    try:
        # 1. Get file from uploads
        file_data = get_file_from_minio(MINIO_BUCKET_UPLOADS, file_id)

        # 2. Process
        processed_buffer = process_presentation(file_data, request)

        # 3. Upload to processed bucket
        upload_file_to_minio(MINIO_BUCKET_PROCESSED, file_id, processed_buffer, "application/vnd.openxmlformats-officedocument.presentationml.presentation")

        return ProcessResponse(file_id=file_id, processed_file_id=file_id, message="File processed and saved successfully")

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Processing error: {str(e)}")


@router.get("/list-processed", response_model=ListProcessedResponse)
async def list_processed():
    files = list_files_in_bucket(MINIO_BUCKET_PROCESSED)
    return ListProcessedResponse(files=files)
