import io
from typing import Generator
from pptx import Presentation
from timeline import set_sidebar_timeline, move_elements_to_right, set_morph_transitions, Configurations
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
            transition_duration=self.request.transition_duration,
            apply_morph_transition=self.request.apply_morph_transition,
        )

        yield {"stage": "moving", "progress": 10, "message": "repositioning elements..."}
        yield from self._move_elements_progressive()

        yield {"stage": "sidebar", "progress": 50, "message": "creating sidebar timeline..."}
        yield from self._set_sidebar_progressive()

        if self.request.apply_morph_transition:
            yield {"stage": "transitions", "progress": 85, "message": "applying morph transitions..."}
            yield from self._set_transitions_progressive()

        yield {"stage": "saving", "progress": 95, "message": "saving presentation..."}

        output_buffer = io.BytesIO()
        self.prs.save(output_buffer)
        output_buffer.seek(0)

        yield {"stage": "complete", "progress": 100, "message": "done"}

        return output_buffer

    def _move_elements_progressive(self) -> Generator[dict, None, None]:
        """move elements with per-slide progress"""
        for i, slide in enumerate(self.prs.slides):
            for shape in slide.shapes:
                original_top = shape.top
                original_left = shape.left
                original_height = shape.height
                original_width = shape.width

                content_space_width = self.prs.slide_width * (1 - self.config.sidebar_width)
                scale_factor = content_space_width / self.prs.slide_width

                new_width = original_width * scale_factor
                new_height = original_height * scale_factor

                new_left = self.prs.slide_width * self.config.sidebar_width + (original_left - original_width / 2) * scale_factor + new_width / 2
                vertical_center_offset = (original_height - new_height) / 2
                new_top = original_top + vertical_center_offset

                shape.left, shape.width = int(new_left), int(new_width)
                shape.top, shape.height = int(new_top), int(new_height)

            progress = 10 + int((i + 1) / self.total_slides * 35)
            yield {"stage": "moving", "progress": progress, "message": f"repositioning slide {i + 1}/{self.total_slides}"}

    def _set_sidebar_progressive(self) -> Generator[dict, None, None]:
        """set sidebar timeline with progress updates"""
        set_sidebar_timeline(ppt=self.prs, tags=self.request.tags, config=self.config)

        for i in range(self.total_slides):
            progress = 50 + int((i + 1) / self.total_slides * 30)
            yield {"stage": "sidebar", "progress": progress, "message": f"styling slide {i + 1}/{self.total_slides}"}

    def _set_transitions_progressive(self) -> Generator[dict, None, None]:
        """apply transitions with progress"""
        set_morph_transitions(self.prs, config=self.config)

        for i in range(self.total_slides):
            progress = 85 + int((i + 1) / self.total_slides * 10)
            yield {"stage": "transitions", "progress": progress, "message": f"transitions {i + 1}/{self.total_slides}"}
