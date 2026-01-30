import io
from pptx import Presentation
from pptx.dml.color import RGBColor
from timeline import set_sidebar_timeline, move_elements_to_right, set_morph_transitions, Configurations
from webcolors import hex_to_rgb
from ..apis.schemas import ProcessRequest


def process_presentation(file_data: io.BytesIO, request: ProcessRequest) -> io.BytesIO:
    # 1. load presentation
    prs = Presentation(file_data)

    # 2. configure
    config = Configurations(
        sidebar_width=request.sidebar_width,
        sidebar_item_height=request.sidebar_item_height,
        sidebar_init_font_size=request.sidebar_init_font_size,
        sidebar_color=RGBColor(*hex_to_rgb(request.sidebar_color_hex)),
        indicator_color=RGBColor(*hex_to_rgb(request.indicator_color_hex)),
        sidebar_item_font_color=RGBColor(*hex_to_rgb(request.sidebar_item_font_color_hex)),
        transition_duration=request.transition_duration,
        apply_morph_transition=request.apply_morph_transition,
    )

    # 3. process
    move_elements_to_right(prs, config=config)
    set_sidebar_timeline(ppt=prs, tags=request.tags, config=config)
    if request.apply_morph_transition:
        set_morph_transitions(prs, config=config)

    # 4. save to buffer
    output_buffer = io.BytesIO()
    prs.save(output_buffer)
    output_buffer.seek(0)

    return output_buffer
