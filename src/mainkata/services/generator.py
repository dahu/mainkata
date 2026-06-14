#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict

from mainkata.backgrounds.images import build_background_pool
from mainkata.config.style_config import (load_style_config,
                                          resolve_title_slide_style,
                                          resolve_vocab_slide_style)
from mainkata.domain.selection import random_sets
from mainkata.domain.types import BackgroundMode, PrimarySide
from mainkata.domain.validation import (validate_background_options,
                                        validate_generation_options,
                                        validate_visual_options)
from mainkata.io.paths import resolve_csv_path, resolve_output_path
from mainkata.io.selected_terms_csv import write_selected_terms_csv
from mainkata.io.vocab_csv import read_vocab_csv
from mainkata.pptx.deck_builder import build_pptx


def _apply_style_overrides(
    style_config: Dict[str, Any],
    title_slide_overlay_transparency: float | None = None,
    vocab_slide_overlay_transparency: float | None = None,
    show_title_card: bool | None = None,
    title_card_transparency: float | None = None,
    show_vocab_card: bool | None = None,
    vocab_card_transparency: float | None = None,
) -> tuple[Dict[str, Any], Dict[str, Any]]:
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

    return title_style, vocab_style


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
    # Validate high-level options
    validate_generation_options(set_count, set_size, primary_side)
    validate_background_options(
        background_dir=background_dir,
        background_mode=background_mode,
        background_image_number=background_image_number,
        background_cycle_start=background_cycle_start,
        background_cycle_end=background_cycle_end,
    )

    # Resolve paths and load style config
    csv_path = resolve_csv_path(csv_file)
    output_path = resolve_output_path(csv_path, output)
    style_config = load_style_config(style_config_file)

    # Resolve styles, applying any overrides
    title_style, vocab_style = _apply_style_overrides(
        style_config=style_config,
        title_slide_overlay_transparency=title_slide_overlay_transparency,
        vocab_slide_overlay_transparency=vocab_slide_overlay_transparency,
        show_title_card=show_title_card,
        title_card_transparency=title_card_transparency,
        show_vocab_card=show_vocab_card,
        vocab_card_transparency=vocab_card_transparency,
    )

    # Validate final visual values
    validate_visual_options(
        title_slide_overlay_transparency=title_style["overlay_transparency"],
        vocab_slide_overlay_transparency=vocab_style["overlay_transparency"],
        title_card_transparency=title_style["card_transparency"],
        vocab_card_transparency=vocab_style["card_transparency"],
    )

    # Prepare vocab sets
    vocab = read_vocab_csv(csv_path, min_rows=set_size)
    sets = random_sets(vocab, set_count=set_count, set_size=set_size, seed=seed)

    # Prepare background pool
    bg_pool = build_background_pool(
        background_dir=background_dir,
        background_mode=background_mode,
        background_image_number=background_image_number,
        background_cycle_start=background_cycle_start,
        background_cycle_end=background_cycle_end,
    )

    # Render deck
    output_path, selected_rows = build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        labels=style_config["labels"],
        sets=sets,
        primary_side=primary_side,
        show_alternate=show_alternate,
        title_style=title_style,
        vocab_style=vocab_style,
        bg_pool=bg_pool,
    )

    # Optional export of selected terms
    csv_out = None
    if export_selected_terms:
        csv_out = write_selected_terms_csv(output_path, selected_rows)

    return output_path, csv_out
