from __future__ import annotations

import csv
from copy import deepcopy
from pathlib import Path

from pptx import Presentation

from mainkata.config.style_config import (DEFAULT_STYLE_CONFIG,
                                          resolve_title_slide_style,
                                          resolve_vocab_slide_style)
from mainkata.domain.options import (BackgroundOptions, GenerationOptions,
                                     VisualOptions)
from mainkata.pptx.deck_builder import build_pptx


def all_slide_text(prs: Presentation) -> list[str]:
    texts: list[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                texts.append(shape.text)
    return texts


def csv_rows(path: Path) -> list[list[str]]:
    with path.open("r", encoding="utf-8", newline="") as f:
        return list(csv.reader(f))


def make_style_config() -> dict:
    return deepcopy(DEFAULT_STYLE_CONFIG)


def make_styles() -> tuple[dict, dict, dict]:
    style_config = make_style_config()
    labels = style_config["labels"]
    title_style = resolve_title_slide_style(style_config)
    vocab_style = resolve_vocab_slide_style(style_config)
    return labels, title_style, vocab_style


def make_generation(
    *,
    set_count: int = 1,
    set_size: int = 2,
    seed: int = 42,
    primary_side: str = "term",
    show_alternate: bool = True,
    export_selected_terms: bool = False,
) -> GenerationOptions:
    return GenerationOptions(
        set_count=set_count,
        set_size=set_size,
        seed=seed,
        primary_side=primary_side,
        show_alternate=show_alternate,
        export_selected_terms=export_selected_terms,
    )


def make_background() -> BackgroundOptions:
    return BackgroundOptions()


def make_visual() -> VisualOptions:
    return VisualOptions()


def make_sets() -> list[list[tuple[str, str]]]:
    return [
        [
            ("cat", "a small domesticated feline"),
            ("dog", "a domesticated canine"),
        ]
    ]


def make_multi_sets() -> list[list[tuple[str, str]]]:
    return [
        [
            ("cat", "a small domesticated feline"),
            ("dog", "a domesticated canine"),
        ],
        [
            ("bird", "a feathered animal"),
            ("fish", "an aquatic vertebrate"),
        ],
    ]


def test_build_pptx_creates_output_file_and_expected_slide_count(tmp_path: Path):
    output_path = tmp_path / "deck.pptx"
    csv_path = tmp_path / "animals.csv"
    csv_path.write_text("placeholder", encoding="utf-8")

    labels, title_style, vocab_style = make_styles()

    pptx_out, csv_out = build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=make_multi_sets(),
        labels=labels,
        title_style=title_style,
        vocab_style=vocab_style,
        generation=make_generation(set_count=2, set_size=2),
        background=make_background(),
        visual=make_visual(),
        bg_pool=[],
    )

    assert pptx_out == output_path
    assert output_path.exists()
    assert csv_out is None

    prs = Presentation(output_path)
    assert len(prs.slides) == 6  # 2 sets * (1 title + 2 vocab)


def test_build_pptx_adds_one_title_slide_per_set(tmp_path: Path):
    output_path = tmp_path / "deck.pptx"
    csv_path = tmp_path / "animals.csv"
    csv_path.write_text("placeholder", encoding="utf-8")

    labels, title_style, vocab_style = make_styles()

    build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=make_multi_sets(),
        labels=labels,
        title_style=title_style,
        vocab_style=vocab_style,
        generation=make_generation(set_count=2, set_size=2),
        background=make_background(),
        visual=make_visual(),
        bg_pool=[],
    )

    prs = Presentation(output_path)
    texts = all_slide_text(prs)

    assert any("SET 1" in text for text in texts)
    assert any("SET 2" in text for text in texts)


def test_build_pptx_exports_selected_terms_csv_when_requested(tmp_path: Path):
    output_path = tmp_path / "deck.pptx"
    csv_path = tmp_path / "animals.csv"
    csv_path.write_text("placeholder", encoding="utf-8")

    labels, title_style, vocab_style = make_styles()

    pptx_out, csv_out = build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=make_sets(),
        labels=labels,
        title_style=title_style,
        vocab_style=vocab_style,
        generation=make_generation(
            set_count=1,
            set_size=2,
            export_selected_terms=True,
        ),
        background=make_background(),
        visual=make_visual(),
        bg_pool=[],
    )

    assert pptx_out.exists()
    assert csv_out is not None
    assert csv_out.exists()

    rows = csv_rows(csv_out)
    assert rows[0] == ["set_number", "term", "definition"]
    assert rows[1] == ["1", "cat", "a small domesticated feline"]
    assert rows[2] == ["1", "dog", "a domesticated canine"]


def test_build_pptx_uses_term_as_primary_and_definition_as_secondary(tmp_path: Path):
    output_path = tmp_path / "deck.pptx"
    csv_path = tmp_path / "animals.csv"
    csv_path.write_text("placeholder", encoding="utf-8")

    labels, title_style, vocab_style = make_styles()

    build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=[[("cat", "a small domesticated feline")]],
        labels=labels,
        title_style=title_style,
        vocab_style=vocab_style,
        generation=make_generation(
            set_count=1,
            set_size=1,
            primary_side="term",
            show_alternate=True,
        ),
        background=make_background(),
        visual=make_visual(),
        bg_pool=[],
    )

    prs = Presentation(output_path)
    texts = all_slide_text(prs)

    assert any("cat" in text for text in texts)
    assert any("a small domesticated feline" in text for text in texts)


def test_build_pptx_uses_definition_as_primary_and_term_as_secondary(tmp_path: Path):
    output_path = tmp_path / "deck.pptx"
    csv_path = tmp_path / "animals.csv"
    csv_path.write_text("placeholder", encoding="utf-8")

    labels, title_style, vocab_style = make_styles()

    build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=[[("cat", "a small domesticated feline")]],
        labels=labels,
        title_style=title_style,
        vocab_style=vocab_style,
        generation=make_generation(
            set_count=1,
            set_size=1,
            primary_side="definition",
            show_alternate=True,
        ),
        background=make_background(),
        visual=make_visual(),
        bg_pool=[],
    )

    prs = Presentation(output_path)
    texts = all_slide_text(prs)

    assert any("a small domesticated feline" in text for text in texts)
    assert any("cat" in text for text in texts)


def test_build_pptx_omits_secondary_text_when_show_alternate_is_false(tmp_path: Path):
    output_path = tmp_path / "deck.pptx"
    csv_path = tmp_path / "animals.csv"
    csv_path.write_text("placeholder", encoding="utf-8")

    labels, title_style, vocab_style = make_styles()

    build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=[[("cat", "a small domesticated feline")]],
        labels=labels,
        title_style=title_style,
        vocab_style=vocab_style,
        generation=make_generation(
            set_count=1,
            set_size=1,
            primary_side="term",
            show_alternate=False,
        ),
        background=make_background(),
        visual=make_visual(),
        bg_pool=[],
    )

    prs = Presentation(output_path)
    texts = all_slide_text(prs)

    assert any("cat" in text for text in texts)
    assert not any("a small domesticated feline" in text for text in texts)


def test_build_pptx_uses_csv_filename_on_title_slide(tmp_path: Path):
    output_path = tmp_path / "deck.pptx"
    csv_path = tmp_path / "ocean_animals.csv"
    csv_path.write_text("placeholder", encoding="utf-8")

    labels, title_style, vocab_style = make_styles()

    build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=make_sets(),
        labels=labels,
        title_style=title_style,
        vocab_style=vocab_style,
        generation=make_generation(set_count=1, set_size=2),
        background=make_background(),
        visual=make_visual(),
        bg_pool=[],
    )

    prs = Presentation(output_path)
    texts = all_slide_text(prs)

    assert any("ocean_animals.csv" in text for text in texts)
