from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple

from pptx import Presentation
from pptx.util import Inches

from mainkata.backgrounds.images import resolve_background_image
from mainkata.domain.types import PrimarySide
from mainkata.pptx.slides import add_title_slide, add_vocab_slide

VocabPair = Tuple[str, str]
VocabSet = List[VocabPair]
SelectedRow = Tuple[int, str, str]


def _build_source_name(csv_path: Path) -> str:
    return csv_path.stem.replace("_", " ").replace("-", " ").title()


def _resolve_slide_text(
    term: str,
    definition: str,
    primary_side: PrimarySide,
    show_alternate: bool,
) -> Tuple[str, str | None]:
    if primary_side == "term":
        primary_text = term
        secondary_text = definition if show_alternate else None
    else:
        primary_text = definition
        secondary_text = term if show_alternate else None

    return primary_text, secondary_text


def build_pptx(
    csv_path: Path,
    output_path: Path,
    labels: Dict[str, str],
    sets: List[VocabSet],
    primary_side: PrimarySide,
    show_alternate: bool,
    title_style: Dict[str, object],
    vocab_style: Dict[str, object],
    bg_pool: List[Path],
) -> Tuple[Path, List[SelectedRow]]:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    source_name = _build_source_name(csv_path)
    selected_rows: List[SelectedRow] = []
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
            primary_text, secondary_text = _resolve_slide_text(
                term=term,
                definition=definition,
                primary_side=primary_side,
                show_alternate=show_alternate,
            )

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

            selected_rows.append((set_number, term, definition))
            generated_slide_index += 1

    prs.save(output_path)
    return output_path, selected_rows
