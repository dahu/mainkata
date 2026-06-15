from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

from mainkata.domain.options import BackgroundOptions, GenerationOptions
from mainkata.pptx.backgrounds import (add_default_background,
                                       add_image_background, add_soft_overlay,
                                       apply_slide_background)
from mainkata.pptx.deck_builder import build_pptx
from mainkata.pptx.slides import add_title_slide, add_vocab_slide
from mainkata.pptx.theme import (apply_font, fit_font_size,
                                 resolve_vocab_primary_font,
                                 set_shape_fill_transparency)


def make_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def make_colors() -> dict[str, RGBColor]:
    return {
        "bg": RGBColor.from_string("F0F7FF"),
        "blue": RGBColor.from_string("2563EB"),
        "text": RGBColor.from_string("0F172A"),
        "subtext": RGBColor.from_string("475569"),
        "coral": RGBColor.from_string("F97066"),
        "white": RGBColor.from_string("FFFFFF"),
    }


def make_title_style() -> dict:
    return {
        "colors": make_colors(),
        "pill_font": {"name": "Aptos", "size": 16, "bold": True},
        "main_font": {"name": "Aptos Display", "size": 28, "bold": True},
        "section_font": {"name": "Aptos Display", "size": 24, "bold": True},
        "body_font": {"name": "Aptos", "size": 16, "bold": False},
        "overlay_transparency": 0.22,
        "show_card": True,
        "card_transparency": 0.18,
    }


def make_vocab_style() -> dict:
    return {
        "colors": make_colors(),
        "primary_font": {"name": "Aptos Display", "size": 24, "bold": True},
        "secondary_font": {"name": "Aptos", "size": 20, "bold": False},
        "overlay_transparency": 0.22,
        "show_card": True,
        "card_transparency": 0.18,
    }


def write_valid_image(path: Path, color=(255, 0, 0)) -> None:
    from PIL import Image

    image = Image.new("RGB", (16, 16), color=color)
    image.save(path)


# ----------------------------
# theme.py
# ----------------------------


@pytest.mark.parametrize(
    "transparency",
    [-0.1, 1.1],
)
def test_set_shape_fill_transparency_rejects_out_of_range_values(
    transparency: float,
) -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(
        1,
        Inches(1),
        Inches(1),
        Inches(2),
        Inches(2),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")

    with pytest.raises(ValueError, match=r"transparency must be between 0\.0 and 1\.0"):
        set_shape_fill_transparency(shape, transparency)


def test_set_shape_fill_transparency_adds_alpha_node() -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(
        1,
        Inches(1),
        Inches(1),
        Inches(2),
        Inches(2),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string("FFFFFF")

    set_shape_fill_transparency(shape, 0.25)

    solid_fill = shape._element.spPr.solidFill
    color_node = solid_fill.srgbClr
    alpha_nodes = [child for child in color_node if child.tag.endswith("alpha")]
    assert len(alpha_nodes) == 1
    assert alpha_nodes[0].get("val") == "75000"


def test_apply_font_sets_run_properties() -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    p = box.text_frame.paragraphs[0]
    run = p.add_run()

    apply_font(
        run,
        {"name": "Aptos", "size": 18, "bold": True},
        RGBColor.from_string("112233"),
    )

    assert run.font.name == "Aptos"
    assert run.font.size == Pt(18)
    assert run.font.bold is True
    assert run.font.color.rgb == RGBColor.from_string("112233")


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("short", 34),
        ("abcdefghijk", 30),
        ("a" * 20, 26),
        ("a" * 30, 22),
        ("a" * 40, 20),
    ],
)
def test_fit_font_size_uses_expected_bands(text: str, expected: int) -> None:
    assert fit_font_size(text) == expected


def test_resolve_vocab_primary_font_copies_and_overrides_size() -> None:
    base = {"name": "Aptos Display", "size": 24, "bold": True}

    resolved = resolve_vocab_primary_font(base, "a" * 40)

    assert resolved["name"] == "Aptos Display"
    assert resolved["bold"] is True
    assert resolved["size"] == 20
    assert base["size"] == 24


# ----------------------------
# backgrounds.py
# ----------------------------


def test_add_default_background_adds_two_oval_shapes() -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_default_background(slide, make_colors())

    assert len(slide.shapes) == 2
    assert all(shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE for shape in slide.shapes)


def test_add_soft_overlay_adds_full_slide_rectangle() -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_soft_overlay(prs, slide, transparency=0.3, colors=make_colors())

    assert len(slide.shapes) == 1
    overlay = slide.shapes[0]
    assert overlay.width == prs.slide_width
    assert overlay.height == prs.slide_height


def test_add_image_background_adds_picture_to_slide(tmp_path: Path) -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    image_path = tmp_path / "bg.png"
    write_valid_image(image_path)

    add_image_background(prs, slide, image_path)

    assert len(slide.shapes) == 1
    picture = slide.shapes[0]
    assert picture.shape_type == MSO_SHAPE_TYPE.PICTURE
    assert picture.width == prs.slide_width
    assert picture.height == prs.slide_height


def test_add_image_background_wraps_picture_load_failures(tmp_path: Path) -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bad_path = tmp_path / "missing.png"

    with pytest.raises(
        RuntimeError,
        match=r"Failed to load background image:",
    ):
        add_image_background(prs, slide, bad_path)


def test_apply_slide_background_uses_default_when_no_image() -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    apply_slide_background(prs, slide, make_colors(), bg_image=None)

    assert len(slide.shapes) == 2


def test_apply_slide_background_uses_image_when_provided(tmp_path: Path) -> None:
    prs = make_prs()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    image_path = tmp_path / "bg.png"
    write_valid_image(image_path)

    apply_slide_background(prs, slide, make_colors(), bg_image=image_path)

    assert len(slide.shapes) == 1
    assert slide.shapes[0].shape_type == MSO_SHAPE_TYPE.PICTURE


# ----------------------------
# slides.py
# ----------------------------


def test_add_title_slide_with_card_creates_expected_text(tmp_path: Path) -> None:
    prs = make_prs()
    style = make_title_style()

    add_title_slide(
        prs,
        set_label="Set 1",
        section_title="Science Vocabulary",
        source_name="terms.csv",
        labels={
            "game_title": "Vocabulary Games",
            "source_prefix": "Source",
            "set_prefix": "Set",
            "vocabulary_suffix": "Vocabulary",
        },
        style=style,
        bg_image=None,
    )

    assert len(prs.slides) == 1
    slide = prs.slides[0]

    texts = [
        shape.text for shape in slide.shapes if getattr(shape, "has_text_frame", False)
    ]
    combined = "\n".join(texts)

    assert "SET 1" in combined
    assert "Vocabulary Games" in combined
    assert "Science Vocabulary" in combined
    assert "Source terms.csv" in combined


def test_add_title_slide_with_background_image_adds_overlay_and_text(
    tmp_path: Path,
) -> None:
    prs = make_prs()
    style = make_title_style()
    image_path = tmp_path / "bg.png"
    write_valid_image(image_path)

    add_title_slide(
        prs,
        set_label="Set 2",
        section_title="History Vocabulary",
        source_name="history.csv",
        labels={
            "game_title": "Vocabulary Games",
            "source_prefix": "Source",
            "set_prefix": "Set",
            "vocabulary_suffix": "Vocabulary",
        },
        style=style,
        bg_image=image_path,
    )

    slide = prs.slides[0]
    texts = [
        shape.text for shape in slide.shapes if getattr(shape, "has_text_frame", False)
    ]
    combined = "\n".join(texts)

    assert "SET 2" in combined
    assert "History Vocabulary" in combined
    assert "Source history.csv" in combined
    assert len(slide.shapes) >= 4


def test_add_vocab_slide_with_card_and_secondary_text_creates_expected_content() -> None:
    prs = make_prs()
    style = make_vocab_style()

    add_vocab_slide(
        prs,
        primary_text="Algorithm",
        secondary_text="A step-by-step procedure",
        style=style,
        bg_image=None,
    )

    assert len(prs.slides) == 1
    slide = prs.slides[0]
    texts = [
        shape.text for shape in slide.shapes if getattr(shape, "has_text_frame", False)
    ]
    combined = "\n".join(texts)

    assert "Algorithm" in combined
    assert "A step-by-step procedure" in combined


def test_add_vocab_slide_without_secondary_text_only_shows_primary() -> None:
    prs = make_prs()
    style = make_vocab_style()

    add_vocab_slide(
        prs,
        primary_text="Loop",
        secondary_text=None,
        style=style,
        bg_image=None,
    )

    slide = prs.slides[0]
    texts = [
        shape.text for shape in slide.shapes if getattr(shape, "has_text_frame", False)
    ]
    combined = "\n".join(texts)

    assert "Loop" in combined
    assert "A step-by-step procedure" not in combined


def test_add_vocab_slide_without_card_uses_text_only_layout() -> None:
    prs = make_prs()
    style = make_vocab_style()
    style["show_card"] = False

    add_vocab_slide(
        prs,
        primary_text="Variable",
        secondary_text="A named storage location",
        style=style,
        bg_image=None,
    )

    slide = prs.slides[0]
    texts = [
        shape.text for shape in slide.shapes if getattr(shape, "has_text_frame", False)
    ]
    combined = "\n".join(texts)

    assert "Variable" in combined
    assert "A named storage location" in combined


# ----------------------------
# deck_builder.py
# ----------------------------


def test_build_pptx_creates_title_and_vocab_slides(tmp_path: Path) -> None:
    csv_path = tmp_path / "computer_terms.csv"
    csv_path.write_text(
        "Term,Definition\nCPU,Processor\nRAM,Memory\n", encoding="utf-8"
    )
    output_path = tmp_path / "deck.pptx"

    sets = [
        [("CPU", "Processor"), ("RAM", "Memory")],
        [("Disk", "Storage"), ("GPU", "Graphics")],
    ]

    result = build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=sets,
        labels={
            "game_title": "Vocabulary Games",
            "source_prefix": "Source",
            "set_prefix": "Set",
            "vocabulary_suffix": "Vocabulary",
        },
        title_style=make_title_style(),
        vocab_style=make_vocab_style(),
        generation=GenerationOptions(
            set_count=2,
            set_size=2,
            seed=42,
            primary_side="term",
            show_alternate=True,
            export_selected_terms=False,
        ),
        background=BackgroundOptions(),
        bg_pool=[],
    )

    assert result.pptx_path == output_path
    assert result.selected_terms_csv_path is None
    assert output_path.exists()

    prs = Presentation(output_path)
    assert len(prs.slides) == 6


def test_build_pptx_uses_definition_as_primary_side(tmp_path: Path) -> None:
    csv_path = tmp_path / "terms.csv"
    csv_path.write_text("Term,Definition\nCPU,Processor\n", encoding="utf-8")
    output_path = tmp_path / "deck.pptx"

    build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=[[("CPU", "Processor")]],
        labels={
            "game_title": "Vocabulary Games",
            "source_prefix": "Source",
            "set_prefix": "Set",
            "vocabulary_suffix": "Vocabulary",
        },
        title_style=make_title_style(),
        vocab_style=make_vocab_style(),
        generation=GenerationOptions(
            set_count=1,
            set_size=1,
            seed=42,
            primary_side="definition",
            show_alternate=True,
            export_selected_terms=False,
        ),
        background=BackgroundOptions(),
        bg_pool=[],
    )

    prs = Presentation(output_path)
    vocab_slide = prs.slides[1]
    texts = [
        shape.text
        for shape in vocab_slide.shapes
        if getattr(shape, "has_text_frame", False)
    ]
    combined = "\n".join(texts)

    assert "Processor" in combined
    assert "CPU" in combined


def test_build_pptx_hides_alternate_when_requested(tmp_path: Path) -> None:
    csv_path = tmp_path / "terms.csv"
    csv_path.write_text("Term,Definition\nCPU,Processor\n", encoding="utf-8")
    output_path = tmp_path / "deck.pptx"

    build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=[[("CPU", "Processor")]],
        labels={
            "game_title": "Vocabulary Games",
            "source_prefix": "Source",
            "set_prefix": "Set",
            "vocabulary_suffix": "Vocabulary",
        },
        title_style=make_title_style(),
        vocab_style=make_vocab_style(),
        generation=GenerationOptions(
            set_count=1,
            set_size=1,
            seed=42,
            primary_side="term",
            show_alternate=False,
            export_selected_terms=False,
        ),
        background=BackgroundOptions(),
        bg_pool=[],
    )

    prs = Presentation(output_path)
    vocab_slide = prs.slides[1]
    texts = [
        shape.text
        for shape in vocab_slide.shapes
        if getattr(shape, "has_text_frame", False)
    ]
    combined = "\n".join(texts)

    assert "CPU" in combined
    assert "Processor" not in combined


def test_build_pptx_exports_selected_terms_when_enabled(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    csv_path = tmp_path / "terms.csv"
    csv_path.write_text("Term,Definition\nCPU,Processor\n", encoding="utf-8")
    output_path = tmp_path / "deck.pptx"
    selected_csv = tmp_path / "deck_selected.csv"

    captured: dict[str, object] = {}

    def fake_write_selected_terms_csv(path, rows):
        captured["path"] = path
        captured["rows"] = rows
        selected_csv.write_text(
            "Set,Term,Definition\n1,CPU,Processor\n", encoding="utf-8"
        )
        return selected_csv

    monkeypatch.setattr(
        "mainkata.pptx.deck_builder.write_selected_terms_csv",
        fake_write_selected_terms_csv,
    )

    result = build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=[[("CPU", "Processor")]],
        labels={
            "game_title": "Vocabulary Games",
            "source_prefix": "Source",
            "set_prefix": "Set",
            "vocabulary_suffix": "Vocabulary",
        },
        title_style=make_title_style(),
        vocab_style=make_vocab_style(),
        generation=GenerationOptions(
            set_count=1,
            set_size=1,
            seed=42,
            primary_side="term",
            show_alternate=True,
            export_selected_terms=True,
        ),
        background=BackgroundOptions(),
        bg_pool=[],
    )

    assert result.pptx_path == output_path
    assert result.selected_terms_csv_path == selected_csv

    assert captured["path"] == output_path
    assert captured["rows"] == [(1, "CPU", "Processor")]


def test_build_pptx_uses_background_pool_for_each_generated_slide(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    csv_path = tmp_path / "terms.csv"
    csv_path.write_text(
        "Term,Definition\nCPU,Processor\nRAM,Memory\n", encoding="utf-8"
    )
    output_path = tmp_path / "deck.pptx"

    calls: list[tuple[list[Path], int]] = []

    def fake_resolve_background_image(bg_pool, generated_slide_index):
        calls.append((list(bg_pool), generated_slide_index))
        return None

    monkeypatch.setattr(
        "mainkata.pptx.deck_builder.resolve_background_image",
        fake_resolve_background_image,
    )

    build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        sets=[[("CPU", "Processor"), ("RAM", "Memory")]],
        labels={
            "game_title": "Vocabulary Games",
            "source_prefix": "Source",
            "set_prefix": "Set",
            "vocabulary_suffix": "Vocabulary",
        },
        title_style=make_title_style(),
        vocab_style=make_vocab_style(),
        generation=GenerationOptions(
            set_count=1,
            set_size=2,
            seed=42,
            primary_side="term",
            show_alternate=True,
            export_selected_terms=False,
        ),
        background=BackgroundOptions(),
        bg_pool=[Path("bg1.png"), Path("bg2.png")],
    )

    assert calls == [
        ([Path("bg1.png"), Path("bg2.png")], 0),
        ([Path("bg1.png"), Path("bg2.png")], 1),
        ([Path("bg1.png"), Path("bg2.png")], 2),
    ]
