from __future__ import annotations

import csv
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from mainkata.backgrounds import resolve_background_image
from mainkata.domain.options import (BackgroundOptions, GenerationOptions,
                                     VisualOptions)
from mainkata.domain.validation import validate_visual_options

from .slides import add_title_slide, add_vocab_slide


def build_pptx(
    csv_path: Path,
    output_path: Path,
    sets,
    labels,
    title_style,
    vocab_style,
    generation: GenerationOptions,
    background: BackgroundOptions,
    visual: VisualOptions,
    bg_pool,
):
    if visual.title_slide_overlay_transparency is not None:
        title_style["overlay_transparency"] = visual.title_slide_overlay_transparency
    if visual.vocab_slide_overlay_transparency is not None:
        vocab_style["overlay_transparency"] = visual.vocab_slide_overlay_transparency
    if visual.show_title_card is not None:
        title_style["show_card"] = visual.show_title_card
    if visual.title_card_transparency is not None:
        title_style["card_transparency"] = visual.title_card_transparency
    if visual.show_vocab_card is not None:
        vocab_style["show_card"] = visual.show_vocab_card
    if visual.vocab_card_transparency is not None:
        vocab_style["card_transparency"] = visual.vocab_card_transparency

    validate_visual_options(
        title_slide_overlay_transparency=title_style["overlay_transparency"],
        vocab_slide_overlay_transparency=vocab_style["overlay_transparency"],
        title_card_transparency=title_style["card_transparency"],
        vocab_card_transparency=vocab_style["card_transparency"],
    )

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    source_name = csv_path.stem.replace("_", " ").replace("-", " ").title()
    rows = []
    generated_slide_index = 0

    for set_number, terms in enumerate(sets, start=1):
        title_bg = (
            resolve_background_image(bg_pool, generated_slide_index)
            if bg_pool
            else None
        )
        add_title_slide(
            prs,
            f"{labels['set_prefix']} {set_number}",
            f"{source_name} {labels['vocabulary_suffix']}",
            csv_path.name,
            labels=labels,
            style=title_style,
            bg_image=title_bg,
        )
        generated_slide_index += 1

        for term, definition in terms:
            if generation.primary_side == "term":
                primary_text = term
                secondary_text = definition if generation.show_alternate else None
            else:
                primary_text = definition
                secondary_text = term if generation.show_alternate else None

            vocab_bg = (
                resolve_background_image(bg_pool, generated_slide_index)
                if bg_pool
                else None
            )
            add_vocab_slide(
                prs,
                primary_text,
                secondary_text,
                style=vocab_style,
                bg_image=vocab_bg,
            )
            rows.append((set_number, term, definition))
            generated_slide_index += 1

    prs.save(output_path)

    csv_out = None
    if generation.export_selected_terms:
        csv_out = output_path.with_name(output_path.stem + "_selected_terms.csv")
        with csv_out.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["set_number", "term", "definition"])
            writer.writerows(rows)

    return output_path, csv_out
