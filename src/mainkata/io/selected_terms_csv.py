from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable, Tuple

Row = Tuple[int, str, str]


def write_selected_terms_csv(
    output_path: Path,
    rows: Iterable[Row],
) -> Path:
    """
    Write the selected terms to a CSV file beside the PPTX.

    The file is named "<pptx_stem>_selected_terms.csv" and contains
    columns: set_number, term, definition.
    """
    csv_out = output_path.with_name(output_path.stem + "_selected_terms.csv")

    with csv_out.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["set_number", "term", "definition"])
        writer.writerows(rows)

    return csv_out
