#!/usr/bin/env python3
from __future__ import annotations

import csv  # temporary
from pathlib import Path
from typing import Any, Dict

from pptx import Presentation
from pptx.util import Inches

from mainkata.backgrounds.images import (build_background_pool,
                                         resolve_background_image)
from mainkata.config.style_config import (load_style_config,
                                          resolve_title_slide_style,
                                          resolve_vocab_slide_style)
from mainkata.domain.selection import random_sets
from mainkata.domain.types import BackgroundMode, PrimarySide
from mainkata.domain.validation import (validate_background_options,
                                        validate_generation_options,
                                        validate_visual_options)
from mainkata.io.paths import resolve_csv_path, resolve_output_path
from mainkata.io.vocab_csv import read_vocab_csv
from mainkata.pptx.slides import add_title_slide, add_vocab_slide


def build_pptx(
    csv_path: Path,
    output_path: Path,
    style_config: Dict[str, Any],
    set_count: int = 6,
    set_size: int = 10,
    seed: int = 42,
    primary_side: PrimarySide = "term",
    show_alternate: bool = True,
    export_selected_terms: bool = False,
    background_dir: str | Path | None = None,
    background_mode: BackgroundMode = "cycle",
    background_image_number: int | None = None,
    background_cycle_start: int | None = None,
    background_cycle_end: int | None = None,
    title_slide_overlay_transparency: float | None = None,
    vocab_slide_overlay_transparency: float | None = None,
    show_title_card: bool | None = None,
    title_card_transparency: float | None = None,
    show_vocab_card: bool | None = None,
    vocab_card_transparency: float | None = None,
):
    vocab = read_vocab_csv(csv_path, min_rows=set_size)
    sets = random_sets(vocab, set_count=set_count, set_size=set_size, seed=seed)

    bg_pool = build_background_pool(
        background_dir=background_dir,
        background_mode=background_mode,
        background_image_number=background_image_number,
        background_cycle_start=background_cycle_start,
        background_cycle_end=background_cycle_end,
    )

    labels = style_config["labels"]
    title_style = resolve_title_slide_style(style_config)
    vocab_style = resolve_vocab_slide_style(style_config)

    if title_slide_overlay_transparency is not None:
        title_style["overlay_transparency"] = title_slide_overlay_transparency
    if vocab_slide_overlay_transparency is not None:
        vocab_style["overlay_transparency"] = vocab_slide_overlay_transparency
    if show_title_card is not None:
        title_style["show_card"] = show_title_card
    if title_card_transparency is not None:
        title_style["card_transparency"] = title_card_transparency
    if show_vocab_card is not None:
        vocab_style["show_card"] = show_vocab_card
    if vocab_card_transparency is not None:
        vocab_style["card_transparency"] = vocab_card_transparency

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
            if primary_side == "term":
                primary_text = term
                secondary_text = definition if show_alternate else None
            else:
                primary_text = definition
                secondary_text = term if show_alternate else None

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
    if export_selected_terms:
        csv_out = output_path.with_name(output_path.stem + "_selected_terms.csv")
        with csv_out.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["set_number", "term", "definition"])
            writer.writerows(rows)

    return output_path, csv_out


def generate_from_inputs(
    csv_file: str | Path,
    output: str | Path | None = None,
    style_config_file: str | Path | None = None,
    set_count: int = 6,
    set_size: int = 10,
    seed: int = 42,
    primary_side: PrimarySide = "term",
    show_alternate: bool = True,
    export_selected_terms: bool = False,
    background_dir: str | Path | None = None,
    background_mode: BackgroundMode = "cycle",
    background_image_number: int | None = None,
    background_cycle_start: int | None = None,
    background_cycle_end: int | None = None,
    title_slide_overlay_transparency: float | None = None,
    vocab_slide_overlay_transparency: float | None = None,
    show_title_card: bool | None = None,
    title_card_transparency: float | None = None,
    show_vocab_card: bool | None = None,
    vocab_card_transparency: float | None = None,
):
    validate_generation_options(set_count, set_size, primary_side)
    validate_background_options(
        background_dir=background_dir,
        background_mode=background_mode,
        background_image_number=background_image_number,
        background_cycle_start=background_cycle_start,
        background_cycle_end=background_cycle_end,
    )

    csv_path = resolve_csv_path(csv_file)
    output_path = resolve_output_path(csv_path, output)
    style_config = load_style_config(style_config_file)

    return build_pptx(
        csv_path,
        output_path,
        style_config=style_config,
        set_count=set_count,
        set_size=set_size,
        seed=seed,
        primary_side=primary_side,
        show_alternate=show_alternate,
        export_selected_terms=export_selected_terms,
        background_dir=background_dir,
        background_mode=background_mode,
        background_image_number=background_image_number,
        background_cycle_start=background_cycle_start,
        background_cycle_end=background_cycle_end,
        title_slide_overlay_transparency=title_slide_overlay_transparency,
        vocab_slide_overlay_transparency=vocab_slide_overlay_transparency,
        show_title_card=show_title_card,
        title_card_transparency=title_card_transparency,
        show_vocab_card=show_vocab_card,
        vocab_card_transparency=vocab_card_transparency,
    )
