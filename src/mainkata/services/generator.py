from __future__ import annotations

from pathlib import Path

from mainkata.backgrounds import build_background_pool
from mainkata.config import (load_style_config, resolve_title_slide_style,
                             resolve_vocab_slide_style)
from mainkata.domain import (BackgroundOptions, GenerationOptions,
                             VisualOptions, random_sets,
                             validate_background_options,
                             validate_generation_options)
from mainkata.io import read_vocab_csv, resolve_csv_path, resolve_output_path
from mainkata.pptx import build_pptx


def generate_from_inputs(
    csv_file: str | Path,
    output: str | Path | None = None,
    style_config_file: str | Path | None = None,
    generation: GenerationOptions = GenerationOptions(),
    background: BackgroundOptions = BackgroundOptions(),
    visual: VisualOptions = VisualOptions(),
):
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
    title_style = resolve_title_slide_style(style_config)
    vocab_style = resolve_vocab_slide_style(style_config)

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
        visual=visual,
        bg_pool=bg_pool,
    )
