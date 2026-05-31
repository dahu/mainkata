from __future__ import annotations

from pathlib import Path
from typing import Any

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from mainkata.pptx.backgrounds import add_soft_overlay, apply_slide_background
from mainkata.pptx.theme import (apply_font, resolve_vocab_primary_font,
                                 set_shape_fill_transparency)


def add_title_card(slide, colors: dict[str, RGBColor]):
    card = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(0.7),
        Inches(0.9),
        Inches(11.9),
        Inches(4.9),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = colors["white"]
    card.line.color.rgb = colors["blue"]
    card.line.transparency = 0.85
    return card


def add_title_slide(
    prs,
    set_label: str,
    section_title: str,
    source_name: str,
    labels: dict[str, str],
    style: dict[str, Any],
    bg_image: Path | None = None,
):
    colors = style["colors"]

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_slide_background(prs, slide, colors, bg_image)

    if bg_image is not None:
        add_soft_overlay(
            prs,
            slide,
            transparency=style["overlay_transparency"],
            colors=colors,
        )

    text_left = Inches(0.9)
    text_top = Inches(1.45)
    text_width = Inches(11.0)
    text_height = Inches(3.9)

    if style["show_card"]:
        card = add_title_card(slide, colors)
        set_shape_fill_transparency(card, style["card_transparency"])
        text_left = Inches(1.1)
        text_top = Inches(1.55)
        text_width = Inches(10.6)
        text_height = Inches(3.5)

    pill = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(0.9),
        Inches(0.7),
        Inches(1.8),
        Inches(0.45),
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = colors["white"]
    set_shape_fill_transparency(pill, 0.15)
    pill.line.color.rgb = colors["blue"]

    tf = pill.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = set_label.upper()
    apply_font(run, style["pill_font"], colors["blue"])

    box = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = labels["game_title"]
    apply_font(r, style["main_font"], colors["text"])

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run()
    r2.text = section_title
    apply_font(r2, style["section_font"], colors["blue"])

    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.LEFT
    r3 = p3.add_run()
    r3.text = f"{labels['source_prefix']} {source_name}"
    apply_font(r3, style["body_font"], colors["subtext"])


def add_vocab_slide(
    prs,
    primary_text: str,
    secondary_text: str | None,
    style: dict[str, Any],
    bg_image: Path | None = None,
):
    colors = style["colors"]

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_slide_background(prs, slide, colors, bg_image)

    if bg_image is not None:
        add_soft_overlay(
            prs,
            slide,
            transparency=style["overlay_transparency"],
            colors=colors,
        )

    text_left = Inches(1.1)
    text_top = Inches(1.55)
    text_width = Inches(11.1)
    text_height = Inches(4.3)

    if style["show_card"]:
        card = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            Inches(0.7),
            Inches(1.0),
            Inches(11.9),
            Inches(5.4),
        )
        card.fill.solid()
        card.fill.fore_color.rgb = colors["white"]
        set_shape_fill_transparency(card, style["card_transparency"])
        card.line.color.rgb = colors["blue"]
        card.line.transparency = 0.85
    else:
        text_left = Inches(0.9)
        text_top = Inches(1.2)
        text_width = Inches(11.5)
        text_height = Inches(4.9)

    box = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.clear()

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = primary_text
    apply_font(
        r1,
        resolve_vocab_primary_font(style["primary_font"], primary_text),
        colors["text"],
    )

    if secondary_text:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(12)
        r2 = p2.add_run()
        r2.text = secondary_text
        apply_font(r2, style["secondary_font"], colors["subtext"])
