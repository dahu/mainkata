#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from mainkata.services.generator import generate_from_inputs, resolve_output_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate a PowerPoint deck from one Term-Definition CSV file."
    )
    parser.add_argument(
        "csv_file", help="Path to input CSV with headers: Term, Definition"
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output PPTX filename; default is based on the CSV filename",
    )
    parser.add_argument(
        "--sets", type=int, default=6, help="Number of sets to generate (default: 6)"
    )
    parser.add_argument(
        "--set-size",
        type=int,
        default=10,
        help="Slides per set, excluding the title slide (default: 10)",
    )
    parser.add_argument(
        "--seed", type=int, default=42, help="Random seed for repeatable permutations"
    )
    parser.add_argument(
        "--export-selected-terms",
        action="store_true",
        help="Also write a companion CSV listing the selected terms for each set",
    )
    parser.add_argument(
        "--primary-side",
        choices=["term", "definition"],
        default="term",
        help="Which side is shown in large font on each slide (default: term)",
    )
    parser.add_argument(
        "--hide-alternate",
        action="store_true",
        help="Hide the alternate side from the slide instead of showing it in smaller text",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing output file(s) without prompting",
    )
    args = parser.parse_args()

    csv_path = Path(args.csv_file).expanduser().resolve()

    # Determine the intended output path (explicit or default)
    if args.output:
        out_path = Path(args.output).expanduser().resolve()
    else:
        # Use the same default logic as the generator
        out_path = resolve_output_path(csv_path, None)

    # Overwrite guard for both explicit and default outputs
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
            set_count=args.sets,
            set_size=args.set_size,
            seed=args.seed,
            primary_side=args.primary_side,
            show_alternate=not args.hide_alternate,
            export_selected_terms=args.export_selected_terms,
        )
    except Exception as exc:
        parser.exit(1, f"Error: {exc}\n")

    print(f"Created: {pptx_path}")
    if csv_out:
        print(f"Created: {csv_out}")


if __name__ == "__main__":
    main()
