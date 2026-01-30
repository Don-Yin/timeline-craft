"""main api router - combines all sub-routers"""
import logging

from fastapi import APIRouter

from .thumbnail_routes import router as thumbnail_router
from .preview_routes import router as preview_router
from .processing_routes import router as processing_router

logging.basicConfig(level=logging.INFO)

router = APIRouter()

router.include_router(thumbnail_router)
router.include_router(preview_router)
router.include_router(processing_router)
