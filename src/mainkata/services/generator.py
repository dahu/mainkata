#!/usr/bin/env python3
from __future__ import annotations

import csv
import random
from pathlib import Path
from typing import List, Literal, Tuple

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

BG = RGBColor(240, 247, 255)
BLUE = RGBColor(37, 99, 235)
TEXT = RGBColor(15, 23, 42)
SUBTEXT = RGBColor(71, 85, 105)
CORAL = RGBColor(249, 112, 102)
WHITE = RGBColor(255, 255, 255)

VocabPair = Tuple[str, str]
PrimarySide = Literal["term", "definition"]
BackgroundMode = Literal["fixed", "cycle"]

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg"}


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


def resolve_background_dir(background_dir: str | Path) -> Path:
    bg_dir = Path(background_dir).expanduser().resolve()
    if not bg_dir.exists():
        raise FileNotFoundError(f"Background directory not found: {bg_dir}")
    if not bg_dir.is_dir():
        raise ValueError(f"Background path is not a directory: {bg_dir}")
    return bg_dir


def validate_generation_options(
    set_count: int,
    set_size: int,
    primary_side: PrimarySide,
) -> None:
    if set_count < 1:
        raise ValueError("--sets must be at least 1.")
    if set_size < 1:
        raise ValueError("--set-size must be at least 1.")
    if primary_side not in {"term", "definition"}:
        raise ValueError("--primary-side must be either 'term' or 'definition'.")


def validate_background_options(
    background_dir: str | Path | None,
    background_mode: BackgroundMode,
    background_image_number: int | None,
    background_cycle_start: int | None,
    background_cycle_end: int | None,
) -> None:
    if background_dir is None:
        return

    if background_mode not in {"fixed", "cycle"}:
        raise ValueError("--background-mode must be 'fixed' or 'cycle'.")

    if background_mode == "fixed":
        if background_image_number is None or background_image_number < 1:
            raise ValueError(
                "--background-image-number must be >= 1 when "
                "--background-mode=fixed."
            )
        if background_cycle_start is not None or background_cycle_end is not None:
            raise ValueError(
                "--background-cycle-start and --background-cycle-end cannot be "
                "used with --background-mode=fixed."
            )

    if background_mode == "cycle":
        if background_image_number is not None:
            raise ValueError(
                "--background-image-number cannot be used with "
                "--background-mode=cycle."
            )

        one_range_value = (background_cycle_start is None) != (
            background_cycle_end is None
        )
        if one_range_value:
            raise ValueError(
                "--background-cycle-start and --background-cycle-end must be "
                "provided together."
            )

        if (
            background_cycle_start is not None
            and background_cycle_end is not None
            and background_cycle_start < 1
        ):
            raise ValueError("--background-cycle-start must be at least 1.")

        if (
            background_cycle_start is not None
            and background_cycle_end is not None
            and background_cycle_end < 1
        ):
            raise ValueError("--background-cycle-end must be at least 1.")

        if (
            background_cycle_start is not None
            and background_cycle_end is not None
            and background_cycle_start > background_cycle_end
        ):
            raise ValueError(
                "--background-cycle-start cannot be greater than "
                "--background-cycle-end."
            )


def validate_visual_options(
    overlay_transparency: float,
    vocab_card_transparency: float,
) -> None:
    if not 0.0 <= overlay_transparency <= 1.0:
        raise ValueError("--overlay-transparency must be between 0.0 and 1.0.")
    if not 0.0 <= vocab_card_transparency <= 1.0:
        raise ValueError("--vocab-card-transparency must be between 0.0 and 1.0.")


def is_valid_image(path: Path) -> bool:
    try:
        with Image.open(path) as im:
            im.verify()
        return True
    except Exception:
        return False


def list_background_images(background_dir: Path) -> List[Path]:
    candidates = sorted(
        p
        for p in background_dir.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES
    )

    bad: List[Path] = []
    images: List[Path] = []

    for p in candidates:
        if is_valid_image(p):
            images.append(p)
        else:
            bad.append(p)

    if bad:
        bad_list = "\n".join(str(p) for p in bad)
        raise ValueError(
            "The following files in the background directory are not valid PNG/JPG "
            "images or use unsupported encodings:\n\n"
            f"{bad_list}\n\n"
            "Please remove or convert them to PNG or JPG."
        )

    if not images:
        raise ValueError(
            f"No valid PNG/JPG images found in background directory: {background_dir}"
        )

    return images


def select_background_pool(
    bg_images: List[Path],
    background_mode: BackgroundMode,
    background_image_number: int | None = None,
    background_cycle_start: int | None = None,
    background_cycle_end: int | None = None,
) -> List[Path]:
    if len(bg_images) == 1:
        return bg_images

    if background_mode == "fixed":
        assert background_image_number is not None
        if background_image_number > len(bg_images):
            raise ValueError(
                f"Requested background image {background_image_number}, "
                f"but only {len(bg_images)} images were found."
            )
        return [bg_images[background_image_number - 1]]

    if background_mode == "cycle":
        if background_cycle_start is None and background_cycle_end is None:
            return bg_images

        assert background_cycle_start is not None
        assert background_cycle_end is not None

        if background_cycle_end > len(bg_images):
            raise ValueError(
                f"Requested background cycle end {background_cycle_end}, "
                f"but only {len(bg_images)} images were found."
            )

        selected = bg_images[background_cycle_start - 1 : background_cycle_end]
        if not selected:
            raise ValueError("Background cycle range did not select any images.")
        return selected

    raise ValueError(f"Unsupported background mode: {background_mode}")


def resolve_background_image(
    bg_pool: List[Path],
    generated_slide_index: int,
) -> Path:
    if not bg_pool:
        raise ValueError("Background image pool is empty.")
    if len(bg_pool) == 1:
        return bg_pool[0]
    return bg_pool[generated_slide_index % len(bg_pool)]


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
    vocab: List[VocabPair],
    set_count: int = 6,
    set_size: int = 10,
    seed: int = 42,
):
    if len(vocab) < set_size:
        raise ValueError(f"Need at least {set_size} unique items; got {len(vocab)}.")
    rng = random.Random(seed)
    return [rng.sample(vocab, set_size) for _ in range(set_count)]


def set_shape_fill_transparency(shape, transparency: float) -> None:
    if not 0.0 <= transparency <= 1.0:
        raise ValueError("transparency must be between 0.0 and 1.0")

    fill = shape.fill
    fill.solid()

    solid_fill = shape._element.spPr.solidFill
    if solid_fill is None:
        raise ValueError("Shape does not have a solid fill.")

    color_node = solid_fill.srgbClr
    if color_node is None:
        color_node = solid_fill.schemeClr
    if color_node is None:
        raise ValueError("Solid fill color node not found.")

    for child in list(color_node):
        if child.tag.endswith("alpha"):
            color_node.remove(child)

    alpha = OxmlElement("a:alpha")
    alpha.set("val", str(int((1.0 - transparency) * 100000)))
    color_node.append(alpha)


def add_default_background(slide):
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
        set_shape_fill_transparency(shape, 0.82)
        shape.line.fill.background()


def add_image_background(prs, slide, image_path: Path):
    try:
        pic = slide.shapes.add_picture(
            str(image_path),
            0,
            0,
            width=prs.slide_width,
            height=prs.slide_height,
        )
    except Exception as exc:
        raise RuntimeError(f"Failed to load background image: {image_path}") from exc

    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)


def add_soft_overlay(prs, slide, transparency: float = 0.22):
    overlay = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        0,
        0,
        prs.slide_width,
        prs.slide_height,
    )
    overlay.fill.solid()
    overlay.fill.fore_color.rgb = WHITE
    set_shape_fill_transparency(overlay, transparency)
    overlay.line.fill.background()


def apply_slide_background(prs, slide, bg_image: Path | None = None):
    if bg_image is not None:
        add_image_background(prs, slide, bg_image)
    else:
        add_default_background(slide)


def add_title_slide(
    prs,
    set_label: str,
    section_title: str,
    source_name: str,
    bg_image: Path | None = None,
    overlay_transparency: float = 0.22,
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_slide_background(prs, slide, bg_image)

    if bg_image is not None:
        add_soft_overlay(prs, slide, transparency=overlay_transparency)

    pill = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(0.9),
        Inches(0.7),
        Inches(1.8),
        Inches(0.45),
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = WHITE
    set_shape_fill_transparency(pill, 0.15)
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


def add_vocab_slide(
    prs,
    primary_text: str,
    secondary_text: str | None = None,
    bg_image: Path | None = None,
    show_vocab_card: bool = True,
    overlay_transparency: float = 0.22,
    vocab_card_transparency: float = 0.18,
):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_slide_background(prs, slide, bg_image)

    if bg_image is not None:
        add_soft_overlay(prs, slide, transparency=overlay_transparency)

    text_left = Inches(1.1)
    text_top = Inches(1.55)
    text_width = Inches(11.1)
    text_height = Inches(4.3)

    if show_vocab_card:
        card = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            Inches(0.7),
            Inches(1.0),
            Inches(11.9),
            Inches(5.4),
        )
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        set_shape_fill_transparency(card, vocab_card_transparency)
        card.line.color.rgb = BLUE
        card.line.transparency = 0.85

    else:
        # Give text slightly more room without the card
        text_left = Inches(0.9)
        text_top = Inches(1.2)
        text_width = Inches(11.5)
        text_height = Inches(4.9)

    box = slide.shapes.add_textbox(text_left, text_top, text_width, text_height)
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
    background_dir: str | Path | None = None,
    background_mode: BackgroundMode = "cycle",
    background_image_number: int | None = None,
    background_cycle_start: int | None = None,
    background_cycle_end: int | None = None,
    overlay_transparency: float = 0.22,
    show_vocab_card: bool = True,
    vocab_card_transparency: float = 0.18,
):
    vocab = read_vocab_csv(csv_path, min_rows=set_size)
    sets = random_sets(vocab, set_count=set_count, set_size=set_size, seed=seed)

    bg_pool: List[Path] = []
    if background_dir is not None:
        bg_dir = resolve_background_dir(background_dir)
        bg_images = list_background_images(bg_dir)
        bg_pool = select_background_pool(
            bg_images,
            background_mode=background_mode,
            background_image_number=background_image_number,
            background_cycle_start=background_cycle_start,
            background_cycle_end=background_cycle_end,
        )

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    source_name = csv_path.stem.replace("_", " ").replace("-", " ").title()
    rows = []
    generated_slide_index = 0

    for set_number, terms in enumerate(sets, start=1):
        title_bg = (
            resolve_background_image(bg_pool, generated_slide_index)
            if bg_pool
            else None
        )
        add_title_slide(
            prs,
            f"Set {set_number}",
            f"{source_name} Vocabulary",
            csv_path.name,
            bg_image=title_bg,
            overlay_transparency=overlay_transparency,
        )
        generated_slide_index += 1

        for term, definition in terms:
            if primary_side == "term":
                primary_text = term
                secondary_text = definition if show_alternate else None
            else:
                primary_text = definition
                secondary_text = term if show_alternate else None

            vocab_bg = (
                resolve_background_image(bg_pool, generated_slide_index)
                if bg_pool
                else None
            )
            add_vocab_slide(
                prs,
                primary_text,
                secondary_text,
                bg_image=vocab_bg,
                overlay_transparency=overlay_transparency,
                show_vocab_card=show_vocab_card,
                vocab_card_transparency=vocab_card_transparency,
            )
            rows.append((set_number, term, definition))
            generated_slide_index += 1

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
    background_dir: str | Path | None = None,
    background_mode: BackgroundMode = "cycle",
    background_image_number: int | None = None,
    background_cycle_start: int | None = None,
    background_cycle_end: int | None = None,
    overlay_transparency: float = 0.22,
    show_vocab_card: bool = True,
    vocab_card_transparency: float = 0.18,
):
    validate_generation_options(set_count, set_size, primary_side)
    validate_background_options(
        background_dir=background_dir,
        background_mode=background_mode,
        background_image_number=background_image_number,
        background_cycle_start=background_cycle_start,
        background_cycle_end=background_cycle_end,
    )
    validate_visual_options(
        overlay_transparency=overlay_transparency,
        vocab_card_transparency=vocab_card_transparency,
    )

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
        background_dir=background_dir,
        background_mode=background_mode,
        background_image_number=background_image_number,
        background_cycle_start=background_cycle_start,
        background_cycle_end=background_cycle_end,
        overlay_transparency=overlay_transparency,
        show_vocab_card=show_vocab_card,
        vocab_card_transparency=vocab_card_transparency,
    )
