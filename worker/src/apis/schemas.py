from pydantic import BaseModel, Field
from typing import List

_HEX_COLOR = r"^#[0-9A-Fa-f]{6}$"


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
    sidebar_color_hex: str = Field(default="#5A5A5A", pattern=_HEX_COLOR, description="Sidebar fill color as hex (e.g., #5A5A5A)")
    indicator_color_hex: str = Field(default="#111111", pattern=_HEX_COLOR, description="Indicator fill color as hex (e.g., #111111)")
    sidebar_item_font_color_hex: str = Field(default="#FFFFFF", pattern=_HEX_COLOR, description="Sidebar font color as hex (e.g., #FFFFFF)")
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


class PreviewRequest(BaseModel):
    """parameters for generating a preview thumbnail with sidebar applied"""

    tags: List[str] = Field(description="List of section tags, one per slide")
    sidebar_width: float = Field(default=0.12, ge=0.05, le=0.5, description="Sidebar width as fraction of slide width")
    sidebar_item_height: float = Field(default=0.10, ge=0.03, le=0.3, description="Height of each tag item as fraction")
    sidebar_color_hex: str = Field(default="#5A5A5A", pattern=_HEX_COLOR, description="Sidebar fill color as hex (e.g., #5A5A5A)")
    indicator_color_hex: str = Field(default="#111111", pattern=_HEX_COLOR, description="Indicator fill color as hex (e.g., #111111)")
    sidebar_item_font_color_hex: str = Field(default="#FFFFFF", pattern=_HEX_COLOR, description="Sidebar font color as hex (e.g., #FFFFFF)")


class AllThumbnailsResponse(BaseModel):
    """response containing all slide thumbnails in one request"""

    file_id: str = Field(description="File ID")
    thumbnails: List[str] = Field(description="List of base64-encoded PNG images, one per slide")


class AllPreviewsResponse(BaseModel):
    """response containing all preview thumbnails with sidebar applied"""

    file_id: str = Field(description="File ID")
    thumbnails: List[str] = Field(description="List of base64-encoded images with sidebar applied")
    format: str = Field(default="png", description="Image format: 'png' or 'jpeg'")
