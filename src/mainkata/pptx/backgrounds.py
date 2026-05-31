from __future__ import annotations

from pathlib import Path

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.util import Inches

from mainkata.pptx.theme import set_shape_fill_transparency


def add_default_background(slide, colors: dict[str, RGBColor]):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colors["bg"]

    for left, top, width, height, color in [
        (Inches(-0.8), Inches(-0.6), Inches(2.6), Inches(2.2), colors["blue"]),
        (Inches(11.0), Inches(5.8), Inches(2.2), Inches(1.8), colors["coral"]),
    ]:
        shape = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.OVAL, left, top, width, height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        set_shape_fill_transparency(shape, 0.82)
        shape.line.fill.background()


def add_image_background(prs, slide, image_path: Path):
    try:
        pic = slide.shapes.add_picture(
            str(image_path),
            0,
            0,
            width=prs.slide_width,
            height=prs.slide_height,
        )
    except Exception as exc:
        raise RuntimeError(f"Failed to load background image: {image_path}") from exc

    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)


def add_soft_overlay(prs, slide, transparency: float, colors: dict[str, RGBColor]):
    overlay = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        0,
        0,
        prs.slide_width,
        prs.slide_height,
    )
    overlay.fill.solid()
    overlay.fill.fore_color.rgb = colors["white"]
    set_shape_fill_transparency(overlay, transparency)
    overlay.line.fill.background()


def apply_slide_background(
    prs,
    slide,
    colors: dict[str, RGBColor],
    bg_image: Path | None = None,
):
    if bg_image is not None:
        add_image_background(prs, slide, bg_image)
    else:
        add_default_background(slide, colors)
