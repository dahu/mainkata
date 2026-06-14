#!/usr/bin/env python3
from __future__ import annotations

import argparse
import traceback
from pathlib import Path

from mainkata.domain import (BackgroundOptions, GenerationOptions,
                             GenerationResult, VisualOptions)
from mainkata.io import resolve_output_path
from mainkata.services.generator import generate_from_inputs


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate a PowerPoint deck from one Term-Definition CSV file.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Print full tracebacks on error",
    )
    parser.add_argument(
        "csvfile",
        help="Path to the input CSV file with headers Term, Definition",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output PPTX filename; defaults to a name based on the CSV filename",
    )
    parser.add_argument(
        "--style-config",
        help="Optional TOML style configuration file",
    )

    generation_group = parser.add_argument_group("generation options")
    generation_group.add_argument(
        "--sets", type=int, default=6, help="Number of sets to generate"
    )
    generation_group.add_argument(
        "--set-size",
        type=int,
        default=10,
        help="Slides per set, excluding the title slide",
    )
    generation_group.add_argument(
        "--seed", type=int, default=42, help="Random seed for repeatable permutations"
    )
    generation_group.add_argument(
        "--export-selected-terms",
        action="store_true",
        help="Also write a companion CSV listing the selected terms for each set",
    )
    generation_group.add_argument(
        "--primary-side",
        choices=("term", "definition"),
        default="term",
        help="Which side is shown in large font on each slide",
    )
    generation_group.add_argument(
        "--hide-alternate",
        action="store_true",
        help="Hide the alternate side instead of showing it in smaller text",
    )

    background_group = parser.add_argument_group("background options")
    background_group.add_argument(
        "--background-dir",
        help="Optional directory containing background image files (.png, .jpg, .jpeg)",
    )
    background_group.add_argument(
        "--background-mode",
        choices=("fixed", "cycle"),
        default="cycle",
        help="How to use multiple background images: fixed uses one nominated image for all slides; cycle rotates through images",
    )
    background_group.add_argument(
        "--background-image-number",
        type=int,
        help="1-based image number to use when --background-mode=fixed",
    )
    background_group.add_argument(
        "--background-cycle-start",
        type=int,
        help="1-based first image number to use when --background-mode=cycle; must be used together with --background-cycle-end",
    )
    background_group.add_argument(
        "--background-cycle-end",
        type=int,
        help="1-based last image number to use when --background-mode=cycle; must be used together with --background-cycle-start",
    )

    slide_style_group = parser.add_argument_group("slide style overrides")
    slide_style_group.add_argument(
        "--title-slide-overlay-transparency",
        type=float,
        default=None,
        help="Override title-slide overlay transparency (0.0 opaque, 1.0 fully transparent)",
    )
    slide_style_group.add_argument(
        "--vocab-slide-overlay-transparency",
        type=float,
        default=None,
        help="Override vocab-slide overlay transparency (0.0 opaque, 1.0 fully transparent)",
    )
    slide_style_group.add_argument(
        "--hide-title-card",
        action="store_true",
        help="Override title slides to hide the white rounded card",
    )
    slide_style_group.add_argument(
        "--show-title-card",
        action="store_true",
        help="Override title slides to show the white rounded card",
    )
    slide_style_group.add_argument(
        "--title-card-transparency",
        type=float,
        default=None,
        help="Override title-slide card transparency (0.0 opaque, 1.0 fully transparent)",
    )
    slide_style_group.add_argument(
        "--hide-vocab-card",
        action="store_true",
        help="Override vocab slides to hide the white rounded card",
    )
    slide_style_group.add_argument(
        "--show-vocab-card",
        action="store_true",
        help="Override vocab slides to show the white rounded card",
    )
    slide_style_group.add_argument(
        "--vocab-card-transparency",
        type=float,
        default=None,
        help="Override vocab-slide card transparency (0.0 opaque, 1.0 fully transparent)",
    )

    output_group = parser.add_argument_group("output options")
    output_group.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing output files without prompting",
    )

    return parser


def resolve_optional_bool(
    parser: argparse.ArgumentParser,
    show_flag: bool,
    hide_flag: bool,
    *,
    show_name: str,
    hide_name: str,
) -> bool | None:
    if show_flag and hide_flag:
        parser.error(f"{show_name} and {hide_name} cannot be used together.")
    if show_flag:
        return True
    if hide_flag:
        return False
    return None


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    csv_path = Path(args.csvfile).expanduser().resolve()
    output_path = resolve_output_path(csv_path, args.output)

    if output_path.exists() and not args.force:
        parser.exit(
            1,
            f"Error: Output file already exists: {output_path}\nUse --force to overwrite.\n",
        )

    generation = GenerationOptions(
        set_count=args.sets,
        set_size=args.set_size,
        seed=args.seed,
        primary_side=args.primary_side,
        show_alternate=not args.hide_alternate,
        export_selected_terms=args.export_selected_terms,
    )

    background = BackgroundOptions(
        background_dir=args.background_dir,
        background_mode=args.background_mode,
        background_image_number=args.background_image_number,
        background_cycle_start=args.background_cycle_start,
        background_cycle_end=args.background_cycle_end,
    )

    visual = VisualOptions(
        title_slide_overlay_transparency=args.title_slide_overlay_transparency,
        vocab_slide_overlay_transparency=args.vocab_slide_overlay_transparency,
        show_title_card=resolve_optional_bool(
            parser,
            args.show_title_card,
            args.hide_title_card,
            show_name="--show-title-card",
            hide_name="--hide-title-card",
        ),
        title_card_transparency=args.title_card_transparency,
        show_vocab_card=resolve_optional_bool(
            parser,
            args.show_vocab_card,
            args.hide_vocab_card,
            show_name="--show-vocab-card",
            hide_name="--hide-vocab-card",
        ),
        vocab_card_transparency=args.vocab_card_transparency,
    )

    try:
        result = generate_from_inputs(
            csv_file=csv_path,
            output=output_path,
            style_config_file=args.style_config,
            generation=generation,
            background=background,
            visual=visual,
        )
    except Exception as exc:
        if getattr(args, "debug", False):
            traceback.print_exc()
        parser.exit(1, f"Error: {exc}\n")

    print(f"Created: {result.pptx_path}")
    if result.selected_terms_csv_path:
        print(f"Created: {result.selected_terms_csv_path}")


if __name__ == "__main__":
    main()
