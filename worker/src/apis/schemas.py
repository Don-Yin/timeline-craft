from pydantic import BaseModel
from typing import List, Optional


class ThumbnailResponse(BaseModel):
    slide_index: int
    image_base64: str


class ProcessRequest(BaseModel):
    tags: List[str]
    sidebar_width: float = 0.12
    sidebar_item_height: float = 0.10
    sidebar_init_font_size: int = 18
    transition_duration: float = 0.3
    apply_morph_transition: bool = True


class ProcessResponse(BaseModel):
    file_id: str
    processed_file_id: str
    message: str


class ListProcessedResponse(BaseModel):
    files: List[str]


class SlideCountResponse(BaseModel):
    file_id: str
    slide_count: int
