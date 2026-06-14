from __future__ import annotations

from pathlib import Path
from typing import Any

from mainkata.backgrounds import build_background_pool
from mainkata.config import (load_style_config, resolve_title_slide_style,
                             resolve_vocab_slide_style)
from mainkata.domain import (BackgroundOptions, GenerationOptions,
                             GenerationResult, VisualOptions, random_sets,
                             validate_background_options,
                             validate_generation_options,
                             validate_visual_options)
from mainkata.io import read_vocab_csv, resolve_csv_path, resolve_output_path
from mainkata.pptx import build_pptx


def apply_visual_overrides(
    title_style: dict[str, Any],
    vocab_style: dict[str, Any],
    visual: VisualOptions,
) -> tuple[dict[str, Any], dict[str, Any]]:
    final_title_style = dict(title_style)
    final_vocab_style = dict(vocab_style)

    if visual.title_slide_overlay_transparency is not None:
        final_title_style[
            "overlay_transparency"
        ] = visual.title_slide_overlay_transparency
    if visual.vocab_slide_overlay_transparency is not None:
        final_vocab_style[
            "overlay_transparency"
        ] = visual.vocab_slide_overlay_transparency

    if visual.show_title_card is not None:
        final_title_style["show_card"] = visual.show_title_card
    if visual.title_card_transparency is not None:
        final_title_style["card_transparency"] = visual.title_card_transparency

    if visual.show_vocab_card is not None:
        final_vocab_style["show_card"] = visual.show_vocab_card
    if visual.vocab_card_transparency is not None:
        final_vocab_style["card_transparency"] = visual.vocab_card_transparency

    required_title_keys = {"overlay_transparency", "card_transparency"}
    required_vocab_keys = {"overlay_transparency", "card_transparency"}

    if required_title_keys.issubset(final_title_style) and required_vocab_keys.issubset(
        final_vocab_style
    ):
        validate_visual_options(
            title_slide_overlay_transparency=final_title_style["overlay_transparency"],
            vocab_slide_overlay_transparency=final_vocab_style["overlay_transparency"],
            title_card_transparency=final_title_style["card_transparency"],
            vocab_card_transparency=final_vocab_style["card_transparency"],
        )

    return final_title_style, final_vocab_style


def generate_from_inputs(
    csv_file: str | Path,
    output: str | Path | None = None,
    style_config_file: str | Path | None = None,
    generation: GenerationOptions = GenerationOptions(),
    background: BackgroundOptions = BackgroundOptions(),
    visual: VisualOptions = VisualOptions(),
) -> GenerationResult:
    validate_generation_options(
        generation.set_count,
        generation.set_size,
        generation.primary_side,
    )
    validate_background_options(
        background_dir=background.background_dir,
        background_mode=background.background_mode,
        background_image_number=background.background_image_number,
        background_cycle_start=background.background_cycle_start,
        background_cycle_end=background.background_cycle_end,
    )

    csv_path = resolve_csv_path(csv_file)
    output_path = resolve_output_path(csv_path, output)

    style_config = load_style_config(style_config_file)
    base_title_style = resolve_title_slide_style(style_config)
    base_vocab_style = resolve_vocab_slide_style(style_config)
    title_style, vocab_style = apply_visual_overrides(
        base_title_style,
        base_vocab_style,
        visual,
    )

    vocab = read_vocab_csv(csv_path, min_rows=generation.set_size)
    sets = random_sets(
        vocab,
        set_count=generation.set_count,
        set_size=generation.set_size,
        seed=generation.seed,
    )
    bg_pool = build_background_pool(background)

    return build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=sets,
        labels=style_config["labels"],
        title_style=title_style,
        vocab_style=vocab_style,
        generation=generation,
        background=background,
        bg_pool=bg_pool,
    )
