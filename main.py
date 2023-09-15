from pathlib import Path
from pptx import Presentation
from src.move_elements import set_sidebar_timeline, move_elements_to_right
from pptx.dml.color import RGBColor
from webcolors import name_to_rgb, hex_to_rgb

sidebar_width = 0.12
sidebar_item_height = 0.06
sidebar_item_font = "Arial"
sidebar_init_font_size = 16
sidebar_item_font_color = RGBColor(255, 255, 255)
sidebar_color = RGBColor(*hex_to_rgb("#5A5A5A"))
sidebar_color_outline = RGBColor(*hex_to_rgb("#FFFFFF"))
sidebar_transparency = 50000  # 50% transparency, alpha * 100000)

indicator_color = RGBColor(*hex_to_rgb("#111111"))
indicator_transparency = 80000


if __name__ == "__main__":
    path_ppt = Path("out", "tbm_ppt.pptx")
    # tags for each slide in an order; has to be the same length as the number of slides
    tags = [
        "Background",
        "Objectives",
        "Data",
        "Curation",
        "Training",
        "Post Analysis",
        "Post Analysis",
        "Limitations",
        "References",
    ]

    example_ppt = Presentation(path_ppt)

    move_elements_to_right(example_ppt, sidebar_width)
    set_sidebar_timeline(
        example_ppt,
        tags,
        sidebar_width,
        sidebar_transparency,
        sidebar_color,
        sidebar_color_outline,
        sidebar_item_height,
        sidebar_init_font_size,
        sidebar_item_font,
        sidebar_item_font_color,
        indicator_color,
        indicator_transparency,
    )
    example_ppt.save(path_ppt.parent / "output.pptx")
