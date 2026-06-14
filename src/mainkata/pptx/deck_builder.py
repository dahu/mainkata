from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from mainkata.backgrounds import resolve_background_image
from mainkata.domain.options import (BackgroundOptions, GenerationOptions,
                                     GenerationResult)
from mainkata.io.selected_terms_csv import write_selected_terms_csv

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
    bg_pool,
):
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
        csv_out = write_selected_terms_csv(output_path, rows)

    return GenerationResult(
        pptx_path=output_path,
        selected_terms_csv_path=csv_out,
    )
