from __future__ import annotations

from pathlib import Path

from mainkata.domain import BackgroundOptions, GenerationOptions, VisualOptions
from mainkata.services import generator


def test_generate_from_inputs_orchestrates_dependencies_and_returns_outputs(
    monkeypatch, tmp_path: Path
) -> None:
    csv_path = tmp_path / "input.csv"
    output_path = tmp_path / "output.pptx"
    exported_csv_path = tmp_path / "output_selected_terms.csv"

    calls: dict[str, object] = {}

    def fake_validate_generation_options(set_count, set_size, primary_side):
        calls["validate_generation_options"] = (set_count, set_size, primary_side)

    def fake_validate_background_options(**kwargs):
        calls["validate_background_options"] = kwargs

    def fake_resolve_csv_path(csv_file):
        calls["resolve_csv_path"] = csv_file
        return csv_path

    def fake_resolve_output_path(resolved_csv_path, output):
        calls["resolve_output_path"] = (resolved_csv_path, output)
        return output_path

    def fake_load_style_config(style_config_file):
        calls["load_style_config"] = style_config_file
        return {
            "labels": {"set_prefix": "Set", "vocabulary_suffix": "Vocabulary"},
            "styles": {},
            "palettes": {},
        }

    def fake_resolve_title_slide_style(style_config):
        calls["resolve_title_slide_style"] = style_config
        return {
            "overlay_transparency": 0.2,
            "card_transparency": 0.1,
            "show_card": True,
            "style_name": "title",
        }

    def fake_resolve_vocab_slide_style(style_config):
        calls["resolve_vocab_slide_style"] = style_config
        return {
            "overlay_transparency": 0.3,
            "card_transparency": 0.15,
            "show_card": False,
            "style_name": "vocab",
        }

    def fake_read_vocab_csv(path, min_rows):
        calls["read_vocab_csv"] = (path, min_rows)
        return [("inu", "dog"), ("neko", "cat")]

    def fake_random_sets(vocab, set_count, set_size, seed):
        calls["random_sets"] = (vocab, set_count, set_size, seed)
        return [[("inu", "dog")]]

    def fake_build_background_pool(options):
        calls["build_background_pool"] = options
        return [Path("bg1.jpg")]

    def fake_build_pptx(**kwargs):
        calls["build_pptx"] = kwargs
        return output_path, exported_csv_path

    monkeypatch.setattr(
        generator, "validate_generation_options", fake_validate_generation_options
    )
    monkeypatch.setattr(
        generator, "validate_background_options", fake_validate_background_options
    )
    monkeypatch.setattr(generator, "resolve_csv_path", fake_resolve_csv_path)
    monkeypatch.setattr(generator, "resolve_output_path", fake_resolve_output_path)
    monkeypatch.setattr(generator, "load_style_config", fake_load_style_config)
    monkeypatch.setattr(
        generator, "resolve_title_slide_style", fake_resolve_title_slide_style
    )
    monkeypatch.setattr(
        generator, "resolve_vocab_slide_style", fake_resolve_vocab_slide_style
    )
    monkeypatch.setattr(generator, "read_vocab_csv", fake_read_vocab_csv)
    monkeypatch.setattr(generator, "random_sets", fake_random_sets)
    monkeypatch.setattr(generator, "build_background_pool", fake_build_background_pool)
    monkeypatch.setattr(generator, "build_pptx", fake_build_pptx)

    generation = GenerationOptions(
        set_count=3,
        set_size=2,
        seed=99,
        primary_side="definition",
        show_alternate=False,
        export_selected_terms=True,
    )
    background = BackgroundOptions(
        background_dir="backgrounds",
        background_mode="fixed",
        background_image_number=2,
    )
    visual = VisualOptions(
        title_slide_overlay_transparency=0.55,
        vocab_slide_overlay_transparency=0.66,
        show_title_card=False,
        title_card_transparency=0.11,
        show_vocab_card=True,
        vocab_card_transparency=0.22,
    )

    pptx_path, csv_out = generator.generate_from_inputs(
        csv_file="lesson.csv",
        output="deck.pptx",
        style_config_file="style.toml",
        generation=generation,
        background=background,
        visual=visual,
    )

    assert pptx_path == output_path
    assert csv_out == exported_csv_path

    assert calls["validate_generation_options"] == (3, 2, "definition")
    assert calls["validate_background_options"] == {
        "background_dir": "backgrounds",
        "background_mode": "fixed",
        "background_image_number": 2,
        "background_cycle_start": None,
        "background_cycle_end": None,
    }
    assert calls["resolve_csv_path"] == "lesson.csv"
    assert calls["resolve_output_path"] == (csv_path, "deck.pptx")
    assert calls["load_style_config"] == "style.toml"
    assert calls["read_vocab_csv"] == (csv_path, 2)
    assert calls["random_sets"] == ([("inu", "dog"), ("neko", "cat")], 3, 2, 99)
    assert calls["build_background_pool"] == background

    build_kwargs = calls["build_pptx"]
    assert build_kwargs["csv_path"] == csv_path
    assert build_kwargs["output_path"] == output_path
    assert build_kwargs["labels"] == {
        "set_prefix": "Set",
        "vocabulary_suffix": "Vocabulary",
    }
    assert build_kwargs["sets"] == [[("inu", "dog")]]
    assert build_kwargs["generation"] == generation
    assert build_kwargs["background"] == background
    assert build_kwargs["visual"] == visual
    assert build_kwargs["bg_pool"] == [Path("bg1.jpg")]
    assert build_kwargs["title_style"]["style_name"] == "title"
    assert build_kwargs["vocab_style"]["style_name"] == "vocab"


def test_generate_from_inputs_handles_defaults_and_empty_background_pool(
    monkeypatch, tmp_path: Path
) -> None:
    csv_path = tmp_path / "input.csv"
    output_path = tmp_path / "output.pptx"

    calls: dict[str, object] = {}

    monkeypatch.setattr(generator, "validate_generation_options", lambda *a: None)
    monkeypatch.setattr(generator, "validate_background_options", lambda **k: None)
    monkeypatch.setattr(generator, "resolve_csv_path", lambda _: csv_path)
    monkeypatch.setattr(generator, "resolve_output_path", lambda *_: output_path)
    monkeypatch.setattr(
        generator,
        "load_style_config",
        lambda _: {
            "labels": {"set_prefix": "Set", "vocabulary_suffix": "Vocabulary"},
            "styles": {},
            "palettes": {},
        },
    )
    monkeypatch.setattr(
        generator,
        "resolve_title_slide_style",
        lambda _: {
            "overlay_transparency": 0.2,
            "card_transparency": 0.1,
            "show_card": True,
        },
    )
    monkeypatch.setattr(
        generator,
        "resolve_vocab_slide_style",
        lambda _: {
            "overlay_transparency": 0.3,
            "card_transparency": 0.15,
            "show_card": True,
        },
    )
    monkeypatch.setattr(generator, "read_vocab_csv", lambda *a, **k: [("inu", "dog")])
    monkeypatch.setattr(generator, "random_sets", lambda *a, **k: [[("inu", "dog")]])

    def fake_build_background_pool(options):
        calls["build_background_pool"] = options
        return []

    def fake_build_pptx(**kwargs):
        calls["build_pptx"] = kwargs
        return output_path, None

    monkeypatch.setattr(generator, "build_background_pool", fake_build_background_pool)
    monkeypatch.setattr(generator, "build_pptx", fake_build_pptx)

    generation = GenerationOptions(set_count=1, set_size=1)
    background = BackgroundOptions()
    visual = VisualOptions()

    pptx_path, csv_out = generator.generate_from_inputs(
        csv_file="lesson.csv",
        output=None,
        style_config_file=None,
        generation=generation,
        background=background,
        visual=visual,
    )

    assert pptx_path == output_path
    assert csv_out is None
    assert calls["build_background_pool"] == background
    assert calls["build_pptx"]["bg_pool"] == []
    assert calls["build_pptx"]["generation"] == generation
    assert calls["build_pptx"]["background"] == background
    assert calls["build_pptx"]["visual"] == visual
