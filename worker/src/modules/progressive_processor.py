import io
from typing import Generator
from pptx import Presentation
from pptx.dml.color import RGBColor
from timeline import set_sidebar_timeline, move_elements_to_right, set_morph_transitions, Configurations
from webcolors import hex_to_rgb
from ..apis.schemas import ProcessRequest


class ProgressiveProcessor:
    """processes presentations with progress callbacks for real-time updates"""

    def __init__(self, file_data: io.BytesIO, request: ProcessRequest):
        self.file_data = file_data
        self.request = request
        self.prs = None
        self.total_slides = 0
        self.config = None

    def process_with_progress(self) -> Generator[dict, None, io.BytesIO]:
        """generator that yields progress updates and returns the processed file"""
        yield {"stage": "loading", "progress": 0, "message": "loading presentation..."}

        self.prs = Presentation(self.file_data)
        self.total_slides = len(self.prs.slides)

        yield {"stage": "loading", "progress": 5, "message": f"loaded {self.total_slides} slides"}

        self.config = Configurations(
            sidebar_width=self.request.sidebar_width,
            sidebar_item_height=self.request.sidebar_item_height,
            sidebar_init_font_size=self.request.sidebar_init_font_size,
            sidebar_color=RGBColor(*hex_to_rgb(self.request.sidebar_color_hex)),
            indicator_color=RGBColor(*hex_to_rgb(self.request.indicator_color_hex)),
            sidebar_item_font_color=RGBColor(*hex_to_rgb(self.request.sidebar_item_font_color_hex)),
            sidebar_transparency=self.request.sidebar_transparency * 1000,
            transition_duration=self.request.transition_duration,
            apply_morph_transition=self.request.apply_morph_transition,
            vertically_center=self.request.vertically_center,
            rounded_indicator=self.request.rounded_indicator,
            center_text=self.request.center_text,
            compact_indicator=self.request.compact_indicator,
        )

        yield {"stage": "moving", "progress": 10, "message": "repositioning elements..."}
        move_elements_to_right(self.prs, config=self.config)
        yield {"stage": "moving", "progress": 45, "message": "elements repositioned"}

        yield {"stage": "sidebar", "progress": 50, "message": "creating sidebar timeline..."}
        set_sidebar_timeline(ppt=self.prs, tags=self.request.tags, config=self.config)
        yield {"stage": "sidebar", "progress": 80, "message": "sidebar created"}

        if self.request.apply_morph_transition:
            yield {"stage": "transitions", "progress": 85, "message": "applying morph transitions..."}
            set_morph_transitions(self.prs, config=self.config)
            yield {"stage": "transitions", "progress": 95, "message": "transitions applied"}

        yield {"stage": "saving", "progress": 95, "message": "saving presentation..."}

        output_buffer = io.BytesIO()
        self.prs.save(output_buffer)
        output_buffer.seek(0)

        yield {"stage": "complete", "progress": 100, "message": "done"}

        return output_buffer
