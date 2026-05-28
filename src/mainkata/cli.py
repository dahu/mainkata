#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from mainkata.services.generator import (
    generate_from_inputs,
    resolve_output_path,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a PowerPoint deck from one Term-Definition CSV file.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    parser.add_argument(
        "csv_file",
        help="Path to the input CSV file with headers: Term, Definition",
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
        "--sets",
        type=int,
        default=6,
        help="Number of sets to generate",
    )
    generation_group.add_argument(
        "--set-size",
        type=int,
        default=10,
        help="Slides per set, excluding the title slide",
    )
    generation_group.add_argument(
        "--seed",
        type=int,
        default=42,
        help="Random seed for repeatable permutations",
    )
    generation_group.add_argument(
        "--export-selected-terms",
        action="store_true",
        help="Also write a companion CSV listing the selected terms for each set",
    )
    generation_group.add_argument(
        "--primary-side",
        choices=["term", "definition"],
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
        choices=["fixed", "cycle"],
        default="cycle",
        help=(
            "How to use multiple background images: "
            "'fixed' uses one nominated image for all slides; "
            "'cycle' rotates through images"
        ),
    )
    background_group.add_argument(
        "--background-image-number",
        type=int,
        help=(
            "1-based image number to use when --background-mode=fixed; "
            "for example, 1 means the first image in sorted order"
        ),
    )
    background_group.add_argument(
        "--background-cycle-start",
        type=int,
        help=(
            "1-based first image number to use when --background-mode=cycle; "
            "must be used together with --background-cycle-end"
        ),
    )
    background_group.add_argument(
        "--background-cycle-end",
        type=int,
        help=(
            "1-based last image number to use when --background-mode=cycle; "
            "must be used together with --background-cycle-start"
        ),
    )

    slide_style_group = parser.add_argument_group("slide style options")
    slide_style_group.add_argument(
        "--title-slide-overlay-transparency",
        type=float,
        default=0.22,
        help=(
            "Transparency for the title-slide soft overlay when using background images "
            "(0.0 = opaque, 1.0 = fully transparent)"
        ),
    )
    slide_style_group.add_argument(
        "--vocab-slide-overlay-transparency",
        type=float,
        default=0.22,
        help=(
            "Transparency for the vocab-slide soft overlay when using background images "
            "(0.0 = opaque, 1.0 = fully transparent)"
        ),
    )
    slide_style_group.add_argument(
        "--hide-title-card",
        action="store_true",
        help="Do not draw the white rounded card on title slides",
    )
    slide_style_group.add_argument(
        "--title-card-transparency",
        type=float,
        default=0.18,
        help=(
            "Transparency for the title-slide white card "
            "(0.0 = opaque, 1.0 = fully transparent)"
        ),
    )
    slide_style_group.add_argument(
        "--hide-vocab-card",
        action="store_true",
        help="Do not draw the white rounded card on vocab slides",
    )
    slide_style_group.add_argument(
        "--vocab-card-transparency",
        type=float,
        default=0.18,
        help=(
            "Transparency for the vocab-slide white card "
            "(0.0 = opaque, 1.0 = fully transparent)"
        ),
    )

    output_group = parser.add_argument_group("output options")
    output_group.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing output file(s) without prompting",
    )

    args = parser.parse_args()

    csv_path = Path(args.csv_file).expanduser().resolve()

    if args.output:
        out_path = Path(args.output).expanduser().resolve()
    else:
        out_path = resolve_output_path(csv_path, None)

    if out_path.exists() and not args.force:
        parser.exit(
            1,
            f"Error: Output file already exists: {out_path}\n"
            "Use --force to overwrite.\n",
        )

    try:
        pptx_path, csv_out = generate_from_inputs(
            csv_file=csv_path,
            output=out_path,
            style_config_file=args.style_config,
            set_count=args.sets,
            set_size=args.set_size,
            seed=args.seed,
            primary_side=args.primary_side,
            show_alternate=not args.hide_alternate,
            export_selected_terms=args.export_selected_terms,
            background_dir=args.background_dir,
            background_mode=args.background_mode,
            background_image_number=args.background_image_number,
            background_cycle_start=args.background_cycle_start,
            background_cycle_end=args.background_cycle_end,
            title_slide_overlay_transparency=args.title_slide_overlay_transparency,
            vocab_slide_overlay_transparency=args.vocab_slide_overlay_transparency,
            show_title_card=not args.hide_title_card,
            title_card_transparency=args.title_card_transparency,
            show_vocab_card=not args.hide_vocab_card,
            vocab_card_transparency=args.vocab_card_transparency,
        )
    except Exception as exc:
        parser.exit(1, f"Error: {exc}\n")

    print(f"Created: {pptx_path}")
    if csv_out:
        print(f"Created: {csv_out}")


if __name__ == "__main__":
    main()
