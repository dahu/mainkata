#!/usr/bin/env python3
from __future__ import annotations

import csv
import random
from pathlib import Path
from typing import List, Literal, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BG = RGBColor(240, 247, 255)
BLUE = RGBColor(37, 99, 235)
TEXT = RGBColor(15, 23, 42)
SUBTEXT = RGBColor(71, 85, 105)
CORAL = RGBColor(249, 112, 102)
WHITE = RGBColor(255, 255, 255)

VocabPair = Tuple[str, str]
PrimarySide = Literal["term", "definition"]


def resolve_csv_path(csv_file: str | Path) -> Path:
    csv_path = Path(csv_file).expanduser().resolve()
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")
    return csv_path


def resolve_output_path(csv_path: Path, output: str | Path | None = None) -> Path:
    if output:
        output_path = Path(output).expanduser().resolve()
    else:
        output_path = csv_path.with_name(csv_path.stem + "_vocab_sets.pptx")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return output_path


def validate_generation_options(
    set_count: int, set_size: int, primary_side: PrimarySide
) -> None:
    if set_count < 1:
        raise ValueError("--sets must be at least 1.")
    if set_size < 1:
        raise ValueError("--set-size must be at least 1.")
    if primary_side not in {"term", "definition"}:
        raise ValueError("--primary-side must be either 'term' or 'definition'.")


def read_vocab_csv(csv_path: Path, min_rows: int = 10) -> List[VocabPair]:
    rows: List[VocabPair] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("CSV file is empty or missing a header row.")

        field_map = {
            name.strip().lower(): name for name in reader.fieldnames if name is not None
        }
        if "term" not in field_map or "definition" not in field_map:
            raise ValueError(
                "CSV must contain headers for both Term and Definition (case-insensitive)."
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

    # Deduplicate first
    unique: List[VocabPair] = []
    seen = set()
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


def random_sets(
    vocab: List[VocabPair], set_count: int = 6, set_size: int = 10, seed: int = 42
):
    if len(vocab) < set_size:
        raise ValueError(f"Need at least {set_size} unique items; got {len(vocab)}.")
    rng = random.Random(seed)
    return [rng.sample(vocab, set_size) for _ in range(set_count)]


def add_background(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG

    for left, top, width, height, color in [
        (Inches(-0.8), Inches(-0.6), Inches(2.6), Inches(2.2), BLUE),
        (Inches(11.0), Inches(5.8), Inches(2.2), Inches(1.8), CORAL),
    ]:
        shape = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.OVAL, left, top, width, height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.fill.transparency = 0.82
        shape.line.fill.background()


def add_title_slide(prs, set_label: str, section_title: str, source_name: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)

    pill = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(0.9),
        Inches(0.7),
        Inches(1.8),
        Inches(0.45),
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = WHITE
    pill.fill.transparency = 0.15
    pill.line.color.rgb = BLUE
    tf = pill.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = set_label.upper()
    run.font.name = "Aptos"
    run.font.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = BLUE

    box = slide.shapes.add_textbox(Inches(0.9), Inches(1.45), Inches(11.0), Inches(2.6))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "Vocabulary Games"
    r.font.name = "Aptos Display"
    r.font.bold = True
    r.font.size = Pt(28)
    r.font.color.rgb = TEXT

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run()
    r2.text = section_title
    r2.font.name = "Aptos Display"
    r2.font.bold = True
    r2.font.size = Pt(24)
    r2.font.color.rgb = BLUE

    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.LEFT
    r3 = p3.add_run()
    r3.text = f"Source: {source_name}"
    r3.font.name = "Aptos"
    r3.font.size = Pt(16)
    r3.font.color.rgb = SUBTEXT


def fit_font_size(text: str) -> int:
    n = len(text)
    if n <= 10:
        return 34
    if n <= 18:
        return 30
    if n <= 26:
        return 26
    if n <= 36:
        return 22
    return 20


def add_vocab_slide(prs, primary_text: str, secondary_text: str | None = None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide)

    card = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(0.7),
        Inches(1.0),
        Inches(11.9),
        Inches(5.4),
    )
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = BLUE
    card.line.transparency = 0.85

    box = slide.shapes.add_textbox(Inches(1.1), Inches(1.55), Inches(11.1), Inches(4.3))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.clear()

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = primary_text
    r1.font.name = "Aptos Display"
    r1.font.bold = True
    r1.font.size = Pt(fit_font_size(primary_text))
    r1.font.color.rgb = TEXT

    if secondary_text:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(12)
        r2 = p2.add_run()
        r2.text = secondary_text
        r2.font.name = "Aptos"
        r2.font.size = Pt(20)
        r2.font.color.rgb = SUBTEXT


def build_pptx(
    csv_path: Path,
    output_path: Path,
    set_count: int = 6,
    set_size: int = 10,
    seed: int = 42,
    primary_side: PrimarySide = "term",
    show_alternate: bool = True,
    export_selected_terms: bool = False,
):
    vocab = read_vocab_csv(csv_path, min_rows=set_size)
    sets = random_sets(vocab, set_count=set_count, set_size=set_size, seed=seed)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    source_name = csv_path.stem.replace("_", " ").replace("-", " ").title()
    rows = []

    for set_number, terms in enumerate(sets, start=1):
        add_title_slide(
            prs, f"Set {set_number}", f"{source_name} Vocabulary", csv_path.name
        )
        for term, definition in terms:
            if primary_side == "term":
                primary_text = term
                secondary_text = definition if show_alternate else None
            else:
                primary_text = definition
                secondary_text = term if show_alternate else None
            add_vocab_slide(prs, primary_text, secondary_text)
            rows.append((set_number, term, definition))

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
    set_count: int = 6,
    set_size: int = 10,
    seed: int = 42,
    primary_side: PrimarySide = "term",
    show_alternate: bool = True,
    export_selected_terms: bool = False,
):
    validate_generation_options(set_count, set_size, primary_side)
    csv_path = resolve_csv_path(csv_file)
    output_path = resolve_output_path(csv_path, output)
    return build_pptx(
        csv_path,
        output_path,
        set_count=set_count,
        set_size=set_size,
        seed=seed,
        primary_side=primary_side,
        show_alternate=show_alternate,
        export_selected_terms=export_selected_terms,
    )
