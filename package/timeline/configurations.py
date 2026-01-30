from pptx.dml.color import RGBColor
from webcolors import hex_to_rgb


class Configurations:
    """configuration class for timeline sidebar settings"""

    def __init__(self, **kwargs):
        self.sidebar_width = kwargs.get("sidebar_width", 0.12)
        self.sidebar_transparency = kwargs.get("sidebar_transparency", 95000)
        self.sidebar_color = kwargs.get("sidebar_color", RGBColor(*hex_to_rgb("#6366f1")))
        self.sidebar_color_outline = kwargs.get("sidebar_color_outline", RGBColor(*hex_to_rgb("#FFFFFF")))
        self.sidebar_item_height = kwargs.get("sidebar_item_height", 0.06)
        self.sidebar_init_font_size = kwargs.get("sidebar_init_font_size", 13)
        self.sidebar_item_font = kwargs.get("sidebar_item_font", "Arial")
        self.sidebar_item_font_color = kwargs.get("sidebar_item_font_color", RGBColor(255, 255, 255))
        self.indicator_color = kwargs.get("indicator_color", RGBColor(*hex_to_rgb("#111111")))
        self.indicator_transparency = kwargs.get("indicator_transparency", 80000)
        self.transition_duration = kwargs.get("transition_duration", 0.3)
        self.apply_morph_transition = kwargs.get("apply_morph_transition", True)
        self.vertically_center = kwargs.get("vertically_center", True)
        self.rounded_indicator = kwargs.get("rounded_indicator", True)
        self.center_text = kwargs.get("center_text", False)
        self.compact_indicator = kwargs.get("compact_indicator", True)
