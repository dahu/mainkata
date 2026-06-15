from __future__ import annotations

from pathlib import Path

import pytest
from PIL import Image
from pptx.dml.color import RGBColor

from mainkata.backgrounds.images import (build_background_pool,
                                         list_background_images,
                                         resolve_background_image,
                                         select_background_pool)
from mainkata.config.style_config import (DEFAULT_STYLE_CONFIG,
                                          deep_merge_dicts, hex_to_rgb_color,
                                          load_style_config,
                                          resolve_color_palette,
                                          resolve_font_palette,
                                          resolve_style_config_path)
from mainkata.domain.options import (BackgroundOptions, GenerationOptions,
                                     VisualOptions)
from mainkata.domain.validation import validate_background_options
from mainkata.services.generator import (apply_visual_overrides,
                                         generate_from_inputs)


def write_valid_image(path: Path, color=(255, 0, 0)) -> None:
    image = Image.new("RGB", (8, 8), color=color)
    image.save(path)


def write_csv(path: Path, rows: list[tuple[str, str]]) -> None:
    lines = ["Term,Definition"]
    lines.extend(f"{term},{definition}" for term, definition in rows)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# ----------------------------
# validate_background_options
# ----------------------------


def test_validate_background_options_returns_early_when_background_dir_is_none() -> None:
    validate_background_options(
        background_dir=None,
        background_mode="fixed",
        background_image_number=None,
        background_cycle_start=99,
        background_cycle_end=1,
    )


def test_validate_background_options_fixed_valid() -> None:
    validate_background_options(
        background_dir=".",
        background_mode="fixed",
        background_image_number=1,
        background_cycle_start=None,
        background_cycle_end=None,
    )


@pytest.mark.parametrize("image_number", [None, 0])
def test_validate_background_options_fixed_requires_positive_image_number(
    image_number: int | None,
) -> None:
    with pytest.raises(
        ValueError,
        match=r"--background-image-number must be >= 1 when --background-mode=fixed\.",
    ):
        validate_background_options(
            background_dir=".",
            background_mode="fixed",
            background_image_number=image_number,
            background_cycle_start=None,
            background_cycle_end=None,
        )


@pytest.mark.parametrize(
    "cycle_start, cycle_end",
    [
        (1, None),
        (None, 2),
        (1, 2),
    ],
)
def test_validate_background_options_fixed_rejects_cycle_args(
    cycle_start: int | None,
    cycle_end: int | None,
) -> None:
    with pytest.raises(
        ValueError,
        match=r"--background-cycle-start and --background-cycle-end cannot be used with --background-mode=fixed\.",
    ):
        validate_background_options(
            background_dir=".",
            background_mode="fixed",
            background_image_number=1,
            background_cycle_start=cycle_start,
            background_cycle_end=cycle_end,
        )


def test_validate_background_options_cycle_rejects_background_image_number() -> None:
    with pytest.raises(
        ValueError,
        match=r"--background-image-number cannot be used with --background-mode=cycle\.",
    ):
        validate_background_options(
            background_dir=".",
            background_mode="cycle",
            background_image_number=1,
            background_cycle_start=None,
            background_cycle_end=None,
        )


@pytest.mark.parametrize(
    "cycle_start, cycle_end",
    [
        (1, None),
        (None, 2),
    ],
)
def test_validate_background_options_cycle_requires_both_range_values(
    cycle_start: int | None,
    cycle_end: int | None,
) -> None:
    with pytest.raises(
        ValueError,
        match=r"--background-cycle-start and --background-cycle-end must be provided together\.",
    ):
        validate_background_options(
            background_dir=".",
            background_mode="cycle",
            background_image_number=None,
            background_cycle_start=cycle_start,
            background_cycle_end=cycle_end,
        )


def test_validate_background_options_cycle_start_must_be_at_least_one() -> None:
    with pytest.raises(
        ValueError,
        match=r"--background-cycle-start must be at least 1\.",
    ):
        validate_background_options(
            background_dir=".",
            background_mode="cycle",
            background_image_number=None,
            background_cycle_start=0,
            background_cycle_end=1,
        )


def test_validate_background_options_cycle_end_must_be_at_least_one() -> None:
    with pytest.raises(
        ValueError,
        match=r"--background-cycle-end must be at least 1\.",
    ):
        validate_background_options(
            background_dir=".",
            background_mode="cycle",
            background_image_number=None,
            background_cycle_start=1,
            background_cycle_end=0,
        )


def test_validate_background_options_cycle_start_cannot_exceed_end() -> None:
    with pytest.raises(
        ValueError,
        match=r"--background-cycle-start cannot be greater than --background-cycle-end\.",
    ):
        validate_background_options(
            background_dir=".",
            background_mode="cycle",
            background_image_number=None,
            background_cycle_start=3,
            background_cycle_end=2,
        )


# ----------------------------
# list_background_images
# ----------------------------


def test_list_background_images_raises_for_invalid_image_files(tmp_path: Path) -> None:
    good = tmp_path / "good.png"
    bad = tmp_path / "bad.png"
    ignored = tmp_path / "notes.txt"

    write_valid_image(good)
    bad.write_bytes(b"not a real image")
    ignored.write_text("hello", encoding="utf-8")

    with pytest.raises(
        ValueError,
        match=r"The following files in the background directory are not valid PNG/JPG images",
    ) as excinfo:
        list_background_images(tmp_path)

    message = str(excinfo.value)
    assert str(bad) in message
    assert str(good) not in message
    assert str(ignored) not in message


def test_list_background_images_raises_when_no_valid_images_found(
    tmp_path: Path,
) -> None:
    (tmp_path / "notes.txt").write_text("hello", encoding="utf-8")

    with pytest.raises(
        ValueError,
        match=r"No valid PNG/JPG images found in background directory:",
    ):
        list_background_images(tmp_path)


def test_list_background_images_returns_sorted_valid_images(tmp_path: Path) -> None:
    img2 = tmp_path / "b.png"
    img1 = tmp_path / "a.jpg"

    write_valid_image(img2)
    write_valid_image(img1)

    result = list_background_images(tmp_path)

    assert result == [img1, img2]


# ----------------------------
# select_background_pool
# ----------------------------


def test_select_background_pool_returns_single_image_unchanged() -> None:
    images = [Path("only.png")]

    result = select_background_pool(
        images,
        background_mode="fixed",
        background_image_number=1,
    )

    assert result == images


def test_select_background_pool_fixed_happy_path() -> None:
    images = [Path("1.png"), Path("2.png"), Path("3.png")]

    result = select_background_pool(
        images,
        background_mode="fixed",
        background_image_number=2,
    )

    assert result == [Path("2.png")]


def test_select_background_pool_fixed_out_of_range() -> None:
    images = [Path("1.png"), Path("2.png"), Path("3.png")]

    with pytest.raises(
        ValueError,
        match=r"Requested background image 4, but only 3 images were found\.",
    ):
        select_background_pool(
            images,
            background_mode="fixed",
            background_image_number=4,
        )


def test_select_background_pool_cycle_with_no_range_returns_all() -> None:
    images = [Path("1.png"), Path("2.png"), Path("3.png")]

    result = select_background_pool(
        images,
        background_mode="cycle",
        background_cycle_start=None,
        background_cycle_end=None,
    )

    assert result == images


def test_select_background_pool_cycle_with_range_returns_slice() -> None:
    images = [Path("1.png"), Path("2.png"), Path("3.png")]

    result = select_background_pool(
        images,
        background_mode="cycle",
        background_cycle_start=1,
        background_cycle_end=2,
    )

    assert result == [Path("1.png"), Path("2.png")]


def test_select_background_pool_cycle_end_out_of_range() -> None:
    images = [Path("1.png"), Path("2.png"), Path("3.png")]

    with pytest.raises(
        ValueError,
        match=r"Requested background cycle end 4, but only 3 images were found\.",
    ):
        select_background_pool(
            images,
            background_mode="cycle",
            background_cycle_start=1,
            background_cycle_end=4,
        )


def test_select_background_pool_cycle_start_beyond_end_of_list_gives_out_of_range_error() -> None:
    images = [Path("1.png"), Path("2.png"), Path("3.png")]

    with pytest.raises(
        ValueError,
        match=r"Requested background cycle end 4, but only 3 images were found\.",
    ):
        select_background_pool(
            images,
            background_mode="cycle",
            background_cycle_start=4,
            background_cycle_end=4,
        )


# ----------------------------
# resolve_background_image
# ----------------------------


def test_resolve_background_image_raises_for_empty_pool() -> None:
    with pytest.raises(ValueError, match=r"Background image pool is empty\."):
        resolve_background_image([], generated_slide_index=0)


def test_resolve_background_image_returns_single_image_for_any_index() -> None:
    pool = [Path("only.png")]

    assert resolve_background_image(pool, 0) == Path("only.png")
    assert resolve_background_image(pool, 999) == Path("only.png")


@pytest.mark.parametrize(
    ("index", "expected"),
    [
        (0, Path("1.png")),
        (1, Path("2.png")),
        (2, Path("3.png")),
        (3, Path("1.png")),
        (4, Path("2.png")),
        (5, Path("3.png")),
    ],
)
def test_resolve_background_image_cycles_through_pool(
    index: int,
    expected: Path,
) -> None:
    pool = [Path("1.png"), Path("2.png"), Path("3.png")]
    assert resolve_background_image(pool, index) == expected


# ----------------------------
# style config path and load
# ----------------------------


def test_resolve_style_config_path_explicit_missing_file_raises(tmp_path: Path) -> None:
    missing = tmp_path / "missing.toml"

    with pytest.raises(FileNotFoundError, match=r"Style config file not found:"):
        resolve_style_config_path(missing)


def test_resolve_style_config_path_explicit_directory_raises(tmp_path: Path) -> None:
    directory = tmp_path / "configdir"
    directory.mkdir()

    with pytest.raises(ValueError, match=r"Style config path is not a file:"):
        resolve_style_config_path(directory)


def test_load_style_config_returns_default_when_no_file_exists(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))

    loaded = load_style_config()

    assert loaded == DEFAULT_STYLE_CONFIG
    assert loaded is not DEFAULT_STYLE_CONFIG


def test_load_style_config_uses_default_xdg_file_and_deep_merges(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    config_dir = tmp_path / "mainkata"
    config_dir.mkdir()
    config_file = config_dir / "style.toml"
    config_file.write_text(
        """
[styles.title_slide]
overlay_transparency = 0.55

[labels]
game_title = "My Game"
""".strip(),
        encoding="utf-8",
    )

    loaded = load_style_config()

    assert loaded["styles"]["title_slide"]["overlay_transparency"] == 0.55
    assert loaded["labels"]["game_title"] == "My Game"
    assert (
        loaded["styles"]["title_slide"]["card_transparency"]
        == DEFAULT_STYLE_CONFIG["styles"]["title_slide"]["card_transparency"]
    )
    assert (
        loaded["styles"]["vocab_slide"]["overlay_transparency"]
        == DEFAULT_STYLE_CONFIG["styles"]["vocab_slide"]["overlay_transparency"]
    )


def test_deep_merge_dicts_merges_nested_mappings() -> None:
    base = {"a": {"b": 1, "c": 2}, "x": 3}
    override = {"a": {"c": 99}, "y": 4}

    merged = deep_merge_dicts(base, override)

    assert merged == {"a": {"b": 1, "c": 99}, "x": 3, "y": 4}
    assert base == {"a": {"b": 1, "c": 2}, "x": 3}


# ----------------------------
# colors and fonts
# ----------------------------


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("ffffff", RGBColor(255, 255, 255)),
        ("#000000", RGBColor(0, 0, 0)),
        ("Ff00Aa", RGBColor(255, 0, 170)),
    ],
)
def test_hex_to_rgb_color_valid(value: str, expected: RGBColor) -> None:
    assert hex_to_rgb_color(value) == expected


@pytest.mark.parametrize("value", ["#FFF", "#12345", "#1234567", "GGGGGG"])
def test_hex_to_rgb_color_invalid(value: str) -> None:
    with pytest.raises(ValueError, match=r"Invalid hex color value:"):
        hex_to_rgb_color(value)


def test_resolve_color_palette_unknown_name_raises() -> None:
    with pytest.raises(ValueError, match=r"Unknown color palette: no_such_palette"):
        resolve_color_palette(DEFAULT_STYLE_CONFIG, "no_such_palette")


def test_resolve_color_palette_returns_rgb_colors() -> None:
    palette = resolve_color_palette(DEFAULT_STYLE_CONFIG, "default")

    assert palette["bg"] == RGBColor.from_string("F0F7FF")
    assert palette["blue"] == RGBColor.from_string("2563EB")


def test_resolve_font_palette_unknown_name_raises() -> None:
    with pytest.raises(ValueError, match=r"Unknown font palette: no_such_font"):
        resolve_font_palette(DEFAULT_STYLE_CONFIG, "no_such_font")


def test_resolve_font_palette_returns_typed_values() -> None:
    font = resolve_font_palette(DEFAULT_STYLE_CONFIG, "title_main")

    assert font == {
        "name": "Aptos Display",
        "size": 28,
        "bold": True,
    }


# ----------------------------
# apply_visual_overrides
# ----------------------------


def make_base_styles() -> tuple[dict, dict]:
    title_style = {
        "overlay_transparency": 0.22,
        "show_card": True,
        "card_transparency": 0.18,
    }
    vocab_style = {
        "overlay_transparency": 0.22,
        "show_card": True,
        "card_transparency": 0.18,
    }
    return title_style, vocab_style


def test_apply_visual_overrides_with_no_overrides_returns_same_values() -> None:
    title_style, vocab_style = make_base_styles()

    final_title, final_vocab = apply_visual_overrides(
        title_style,
        vocab_style,
        VisualOptions(),
    )

    assert final_title == title_style
    assert final_vocab == vocab_style
    assert final_title is not title_style
    assert final_vocab is not vocab_style


def test_apply_visual_overrides_applies_individual_overrides() -> None:
    title_style, vocab_style = make_base_styles()

    visual = VisualOptions(
        title_slide_overlay_transparency=0.5,
        show_title_card=False,
        vocab_card_transparency=0.4,
    )

    final_title, final_vocab = apply_visual_overrides(
        title_style,
        vocab_style,
        visual,
    )

    assert final_title["overlay_transparency"] == 0.5
    assert final_title["show_card"] is False
    assert final_title["card_transparency"] == 0.18

    assert final_vocab["overlay_transparency"] == 0.22
    assert final_vocab["show_card"] is True
    assert final_vocab["card_transparency"] == 0.4


@pytest.mark.parametrize(
    "visual",
    [
        VisualOptions(title_slide_overlay_transparency=1.5),
        VisualOptions(vocab_slide_overlay_transparency=-0.1),
        VisualOptions(title_card_transparency=1.1),
        VisualOptions(vocab_card_transparency=-0.2),
    ],
)
def test_apply_visual_overrides_validates_transparency_ranges(
    visual: VisualOptions,
) -> None:
    title_style, vocab_style = make_base_styles()

    with pytest.raises(ValueError, match=r"must be between 0\.0 and 1\.0"):
        apply_visual_overrides(title_style, vocab_style, visual)


# ----------------------------
# build_background_pool
# ----------------------------


def test_build_background_pool_returns_empty_list_when_no_background_dir() -> None:
    options = BackgroundOptions(background_dir=None)

    assert build_background_pool(options) == []


def test_build_background_pool_reads_and_selects_images(tmp_path: Path) -> None:
    write_valid_image(tmp_path / "1.png")
    write_valid_image(tmp_path / "2.png")
    write_valid_image(tmp_path / "3.png")

    options = BackgroundOptions(
        background_dir=tmp_path,
        background_mode="fixed",
        background_image_number=2,
    )

    result = build_background_pool(options)

    assert result == [tmp_path / "2.png"]


# ----------------------------
# generate_from_inputs
# ----------------------------


def test_generate_from_inputs_happy_path_with_monkeypatched_dependencies(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_csv(
        csv_path,
        [
            ("T1", "D1"),
            ("T2", "D2"),
            ("T3", "D3"),
            ("T4", "D4"),
        ],
    )

    captured: dict[str, object] = {}

    def fake_build_background_pool(background):
        captured["background"] = background
        return ["bg1", "bg2"]

    def fake_build_pptx(**kwargs):
        captured["build_pptx_kwargs"] = kwargs
        return kwargs["output_path"], None

    monkeypatch.setattr(
        "mainkata.services.generator.build_background_pool", fake_build_background_pool
    )
    monkeypatch.setattr("mainkata.services.generator.build_pptx", fake_build_pptx)

    output = tmp_path / "out.pptx"
    generation = GenerationOptions(
        set_count=2,
        set_size=3,
        seed=7,
        primary_side="definition",
        show_alternate=False,
        export_selected_terms=False,
    )
    visual = VisualOptions(
        title_slide_overlay_transparency=0.4,
        vocab_card_transparency=0.3,
    )

    result = generate_from_inputs(
        csv_file=csv_path,
        output=output,
        generation=generation,
        background=BackgroundOptions(),
        visual=visual,
    )

    assert result == (output, None)

    kwargs = captured["build_pptx_kwargs"]
    assert kwargs["csv_path"] == csv_path
    assert kwargs["output_path"] == output
    assert kwargs["generation"] == generation
    assert kwargs["bg_pool"] == ["bg1", "bg2"]
    assert kwargs["labels"] == DEFAULT_STYLE_CONFIG["labels"]

    assert len(kwargs["sets"]) == 2
    assert all(len(s) == 3 for s in kwargs["sets"])

    assert kwargs["title_style"]["overlay_transparency"] == 0.4
    assert kwargs["vocab_style"]["card_transparency"] == 0.3


def test_generate_from_inputs_rejects_invalid_generation_options(
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("T1", "D1")])

    with pytest.raises(ValueError, match=r"--sets must be at least 1\."):
        generate_from_inputs(
            csv_file=csv_path,
            generation=GenerationOptions(set_count=0, set_size=1),
        )


def test_generate_from_inputs_rejects_invalid_background_options(
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("T1", "D1")])

    with pytest.raises(
        ValueError,
        match=r"--background-image-number must be >= 1 when --background-mode=fixed\.",
    ):
        generate_from_inputs(
            csv_file=csv_path,
            generation=GenerationOptions(set_count=1, set_size=1),
            background=BackgroundOptions(
                background_dir="some-dir",
                background_mode="fixed",
                background_image_number=None,
            ),
        )
