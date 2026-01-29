from pydantic import BaseModel, Field
from typing import List


class ThumbnailResponse(BaseModel):
    """response containing a slide thumbnail"""
    slide_index: int = Field(description="Zero-based slide index")
    image_base64: str = Field(description="Base64-encoded PNG image data")


class ProcessRequest(BaseModel):
    """parameters for processing a presentation with timeline sidebar"""
    tags: List[str] = Field(description="List of section tags, one per slide (e.g., ['intro', 'methods', 'results'])")
    sidebar_width: float = Field(default=0.12, ge=0.05, le=0.5, description="Sidebar width as fraction of slide width (0.05-0.5)")
    sidebar_item_height: float = Field(default=0.10, ge=0.03, le=0.3, description="Height of each tag item as fraction of slide height")
    sidebar_init_font_size: int = Field(default=18, ge=8, le=48, description="Font size for sidebar text in points")
    transition_duration: float = Field(default=0.3, ge=0, le=5, description="Morph transition duration in seconds")
    apply_morph_transition: bool = Field(default=True, description="Whether to apply morph transitions between slides")


class ProcessResponse(BaseModel):
    """response after processing a presentation"""
    file_id: str = Field(description="Original file ID")
    processed_file_id: str = Field(description="Processed file ID (same as original)")
    message: str = Field(description="Status message")


class ListProcessedResponse(BaseModel):
    """list of processed file IDs"""
    files: List[str] = Field(description="Array of processed file IDs")


class SlideCountResponse(BaseModel):
    """response containing slide count"""
    file_id: str = Field(description="File ID")
    slide_count: int = Field(description="Total number of slides in the presentation")
