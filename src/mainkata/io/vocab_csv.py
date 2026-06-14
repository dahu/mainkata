from __future__ import annotations

import csv
from pathlib import Path

from mainkata.domain.types import VocabPair


def read_vocab_csv(csv_path: Path, min_rows: int = 10) -> list[VocabPair]:
    rows: list[VocabPair] = []

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("CSV file is empty or missing a header row.")

        field_map = {
            name.strip().lower(): name for name in reader.fieldnames if name is not None
        }
        if "term" not in field_map or "definition" not in field_map:
            raise ValueError(
                "CSV must contain headers for both Term and Definition "
                "(case-insensitive)."
            )

        term_key = field_map["term"]
        definition_key = field_map["definition"]

        for row in reader:
            term = (row.get(term_key) or "").strip()
            definition = (row.get(definition_key) or "").strip()
            if not term and not definition:
                continue
            if not term or not definition:
                raise ValueError(
                    f"Found incomplete row: Term={term!r}, Definition={definition!r}"
                )
            rows.append((term, definition))

    unique: list[VocabPair] = []
    seen: set[VocabPair] = set()
    for pair in rows:
        if pair not in seen:
            unique.append(pair)
            seen.add(pair)

    if len(unique) < min_rows:
        raise ValueError(
            f"CSV must contain at least {min_rows} unique Term/Definition pairs; "
            f"found {len(unique)}."
        )

    return unique
