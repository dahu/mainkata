from .backgrounds import (add_default_background, add_image_background,
                          add_soft_overlay, apply_slide_background)
from .slides import add_title_card, add_title_slide, add_vocab_slide
from .theme import (apply_font, fit_font_size, resolve_vocab_primary_font,
                    set_shape_fill_transparency)

__all__ = [
    "add_default_background",
    "add_image_background",
    "add_soft_overlay",
    "apply_slide_background",
    "add_title_card",
    "add_title_slide",
    "add_vocab_slide",
    "apply_font",
    "fit_font_size",
    "resolve_vocab_primary_font",
    "set_shape_fill_transparency",
]
