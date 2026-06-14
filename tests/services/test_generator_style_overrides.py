from __future__ import annotations

from pathlib import Path

from mainkata.domain import BackgroundOptions, GenerationOptions, VisualOptions
from mainkata.services import generator


def test_generate_from_inputs_calls_dependencies_and_passes_expected_arguments(
    monkeypatch,
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "input.csv"
    output_path = tmp_path / "output.pptx"

    generation = GenerationOptions(
        set_count=2,
        set_size=3,
        seed=99,
        primary_side="definition",
        show_alternate=False,
        export_selected_terms=True,
    )
    background = BackgroundOptions(
        background_dir="backgrounds",
        background_mode="cycle",
        background_cycle_start=1,
        background_cycle_end=3,
    )
    visual = VisualOptions(
        title_slide_overlay_transparency=0.5,
        show_title_card=False,
        vocab_card_transparency=0.4,
        show_vocab_card=True,
    )

    captured: dict[str, object] = {}

    def fake_validate_generation_options(set_count, set_size, primary_side):
        captured["validate_generation_options"] = (
            set_count,
            set_size,
            primary_side,
        )

    def fake_validate_background_options(**kwargs):
        captured["validate_background_options"] = kwargs

    def fake_resolve_csv_path(value):
        captured["resolve_csv_path"] = value
        return csv_path

    def fake_resolve_output_path(resolved_csv_path, output):
        captured["resolve_output_path"] = (resolved_csv_path, output)
        return output_path

    def fake_load_style_config(style_config_file):
        captured["load_style_config"] = style_config_file
        return {
            "labels": {
                "game_title": "Vocabulary Games",
                "source_prefix": "Source:",
                "set_prefix": "Set",
                "vocabulary_suffix": "Vocabulary",
            }
        }

    def fake_resolve_title_slide_style(style_config):
        captured["resolve_title_slide_style"] = style_config
        return {"kind": "title-style"}

    def fake_resolve_vocab_slide_style(style_config):
        captured["resolve_vocab_slide_style"] = style_config
        return {"kind": "vocab-style"}

    def fake_read_vocab_csv(path, min_rows):
        captured["read_vocab_csv"] = (path, min_rows)
        return [
            ("inu", "dog"),
            ("neko", "cat"),
            ("tori", "bird"),
            ("sakana", "fish"),
        ]

    def fake_random_sets(vocab, set_count, set_size, seed):
        captured["random_sets"] = {
            "vocab": vocab,
            "set_count": set_count,
            "set_size": set_size,
            "seed": seed,
        }
        return [
            [("inu", "dog"), ("neko", "cat"), ("tori", "bird")],
            [("sakana", "fish"), ("inu", "dog"), ("neko", "cat")],
        ]

    def fake_build_background_pool(background_options):
        captured["build_background_pool"] = background_options
        return [Path("bg1.jpg"), Path("bg2.jpg")]

    def fake_build_pptx(**kwargs):
        captured["build_pptx"] = kwargs
        return output_path, None

    monkeypatch.setattr(
        generator,
        "validate_generation_options",
        fake_validate_generation_options,
    )
    monkeypatch.setattr(
        generator,
        "validate_background_options",
        fake_validate_background_options,
    )
    monkeypatch.setattr(generator, "resolve_csv_path", fake_resolve_csv_path)
    monkeypatch.setattr(generator, "resolve_output_path", fake_resolve_output_path)
    monkeypatch.setattr(generator, "load_style_config", fake_load_style_config)
    monkeypatch.setattr(
        generator,
        "resolve_title_slide_style",
        fake_resolve_title_slide_style,
    )
    monkeypatch.setattr(
        generator,
        "resolve_vocab_slide_style",
        fake_resolve_vocab_slide_style,
    )
    monkeypatch.setattr(generator, "read_vocab_csv", fake_read_vocab_csv)
    monkeypatch.setattr(generator, "random_sets", fake_random_sets)
    monkeypatch.setattr(
        generator,
        "build_background_pool",
        fake_build_background_pool,
    )
    monkeypatch.setattr(generator, "build_pptx", fake_build_pptx)

    result = generator.generate_from_inputs(
        csv_file="words.csv",
        output="deck.pptx",
        style_config_file="style.toml",
        generation=generation,
        background=background,
        visual=visual,
    )

    assert result == (output_path, None)

    assert captured["validate_generation_options"] == (2, 3, "definition")
    assert captured["validate_background_options"] == {
        "background_dir": "backgrounds",
        "background_mode": "cycle",
        "background_image_number": None,
        "background_cycle_start": 1,
        "background_cycle_end": 3,
    }

    assert captured["resolve_csv_path"] == "words.csv"
    assert captured["resolve_output_path"] == (csv_path, "deck.pptx")
    assert captured["load_style_config"] == "style.toml"

    assert captured["read_vocab_csv"] == (csv_path, 3)
    assert captured["random_sets"] == {
        "vocab": [
            ("inu", "dog"),
            ("neko", "cat"),
            ("tori", "bird"),
            ("sakana", "fish"),
        ],
        "set_count": 2,
        "set_size": 3,
        "seed": 99,
    }

    assert captured["build_background_pool"] is background

    assert captured["build_pptx"] == {
        "csv_path": csv_path,
        "output_path": output_path,
        "sets": [
            [("inu", "dog"), ("neko", "cat"), ("tori", "bird")],
            [("sakana", "fish"), ("inu", "dog"), ("neko", "cat")],
        ],
        "labels": {
            "game_title": "Vocabulary Games",
            "source_prefix": "Source:",
            "set_prefix": "Set",
            "vocabulary_suffix": "Vocabulary",
        },
        "title_style": {"kind": "title-style"},
        "vocab_style": {"kind": "vocab-style"},
        "generation": generation,
        "background": background,
        "visual": visual,
        "bg_pool": [Path("bg1.jpg"), Path("bg2.jpg")],
    }


def test_generate_from_inputs_uses_generation_set_size_as_csv_min_rows(
    monkeypatch,
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "input.csv"
    output_path = tmp_path / "output.pptx"
    captured: dict[str, object] = {}

    monkeypatch.setattr(generator, "validate_generation_options", lambda *a, **k: None)
    monkeypatch.setattr(generator, "validate_background_options", lambda **k: None)
    monkeypatch.setattr(generator, "resolve_csv_path", lambda _: csv_path)
    monkeypatch.setattr(generator, "resolve_output_path", lambda *_: output_path)
    monkeypatch.setattr(
        generator,
        "load_style_config",
        lambda _: {"labels": {"set_prefix": "Set", "vocabulary_suffix": "Vocabulary"}},
    )
    monkeypatch.setattr(
        generator,
        "resolve_title_slide_style",
        lambda _: {"kind": "title-style"},
    )
    monkeypatch.setattr(
        generator,
        "resolve_vocab_slide_style",
        lambda _: {"kind": "vocab-style"},
    )

    def fake_read_vocab_csv(path, min_rows):
        captured["read_vocab_csv"] = (path, min_rows)
        return [("inu", "dog")] * 5

    monkeypatch.setattr(generator, "read_vocab_csv", fake_read_vocab_csv)
    monkeypatch.setattr(generator, "random_sets", lambda *a, **k: [[("inu", "dog")]])
    monkeypatch.setattr(generator, "build_background_pool", lambda *_: [])
    monkeypatch.setattr(generator, "build_pptx", lambda **kwargs: (output_path, None))

    generation = GenerationOptions(set_size=5)

    generator.generate_from_inputs(
        csv_file="words.csv",
        generation=generation,
    )

    assert captured["read_vocab_csv"] == (csv_path, 5)


def test_generate_from_inputs_passes_visual_options_through_to_builder(
    monkeypatch,
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "input.csv"
    output_path = tmp_path / "output.pptx"
    captured: dict[str, object] = {}

    monkeypatch.setattr(generator, "validate_generation_options", lambda *a, **k: None)
    monkeypatch.setattr(generator, "validate_background_options", lambda **k: None)
    monkeypatch.setattr(generator, "resolve_csv_path", lambda _: csv_path)
    monkeypatch.setattr(generator, "resolve_output_path", lambda *_: output_path)
    monkeypatch.setattr(
        generator,
        "load_style_config",
        lambda _: {"labels": {"set_prefix": "Set", "vocabulary_suffix": "Vocabulary"}},
    )
    monkeypatch.setattr(
        generator,
        "resolve_title_slide_style",
        lambda _: {"kind": "title-style"},
    )
    monkeypatch.setattr(
        generator,
        "resolve_vocab_slide_style",
        lambda _: {"kind": "vocab-style"},
    )
    monkeypatch.setattr(generator, "read_vocab_csv", lambda *a, **k: [("inu", "dog")])
    monkeypatch.setattr(generator, "random_sets", lambda *a, **k: [[("inu", "dog")]])
    monkeypatch.setattr(generator, "build_background_pool", lambda *_: [])

    def fake_build_pptx(**kwargs):
        captured["visual"] = kwargs["visual"]
        return output_path, None

    monkeypatch.setattr(generator, "build_pptx", fake_build_pptx)

    visual = VisualOptions(
        title_slide_overlay_transparency=0.55,
        vocab_slide_overlay_transparency=0.25,
        show_title_card=False,
        title_card_transparency=0.11,
        show_vocab_card=True,
        vocab_card_transparency=0.44,
    )

    generator.generate_from_inputs(
        csv_file="words.csv",
        visual=visual,
    )

    assert captured["visual"] is visual
