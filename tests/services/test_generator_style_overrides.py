from pathlib import Path

from mainkata.services import generator


def test_apply_style_overrides_returns_base_styles_when_no_overrides(
    monkeypatch,
) -> None:
    monkeypatch.setattr(
        generator,
        "resolve_title_slide_style",
        lambda _: {
            "overlay_transparency": 0.2,
            "show_card": True,
            "card_transparency": 0.1,
            "name": "title",
        },
    )
    monkeypatch.setattr(
        generator,
        "resolve_vocab_slide_style",
        lambda _: {
            "overlay_transparency": 0.3,
            "show_card": False,
            "card_transparency": 0.15,
            "name": "vocab",
        },
    )

    title_style, vocab_style = generator._apply_style_overrides({"labels": {}})

    assert title_style == {
        "overlay_transparency": 0.2,
        "show_card": True,
        "card_transparency": 0.1,
        "name": "title",
    }
    assert vocab_style == {
        "overlay_transparency": 0.3,
        "show_card": False,
        "card_transparency": 0.15,
        "name": "vocab",
    }


def test_apply_style_overrides_applies_only_requested_fields(monkeypatch) -> None:
    monkeypatch.setattr(
        generator,
        "resolve_title_slide_style",
        lambda _: {
            "overlay_transparency": 0.2,
            "show_card": True,
            "card_transparency": 0.1,
            "name": "title",
        },
    )
    monkeypatch.setattr(
        generator,
        "resolve_vocab_slide_style",
        lambda _: {
            "overlay_transparency": 0.3,
            "show_card": False,
            "card_transparency": 0.15,
            "name": "vocab",
        },
    )

    title_style, vocab_style = generator._apply_style_overrides(
        {"labels": {}},
        title_slide_overlay_transparency=0.55,
        show_title_card=False,
        vocab_card_transparency=0.42,
        show_vocab_card=True,
    )

    assert title_style["overlay_transparency"] == 0.55
    assert title_style["show_card"] is False
    assert title_style["card_transparency"] == 0.1
    assert title_style["name"] == "title"

    assert vocab_style["overlay_transparency"] == 0.3
    assert vocab_style["show_card"] is True
    assert vocab_style["card_transparency"] == 0.42
    assert vocab_style["name"] == "vocab"


def test_generate_from_inputs_passes_overridden_styles_to_validation_and_renderer(
    monkeypatch,
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "input.csv"
    output_path = tmp_path / "output.pptx"

    captured = {}

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
            "show_card": False,
        },
    )
    monkeypatch.setattr(generator, "read_vocab_csv", lambda *a, **k: [("inu", "dog")])
    monkeypatch.setattr(generator, "random_sets", lambda *a, **k: [[("inu", "dog")]])
    monkeypatch.setattr(generator, "build_background_pool", lambda **k: [])

    def fake_validate_visual_options(**kwargs):
        captured["validated_visuals"] = kwargs

    def fake_build_pptx(**kwargs):
        captured["build_pptx"] = kwargs
        return output_path, [(1, "inu", "dog")]

    monkeypatch.setattr(
        generator, "validate_visual_options", fake_validate_visual_options
    )
    monkeypatch.setattr(generator, "build_pptx", fake_build_pptx)

    generator.generate_from_inputs(
        csv_file="input.csv",
        title_slide_overlay_transparency=0.9,
        show_vocab_card=True,
        vocab_card_transparency=0.8,
    )

    assert captured["validated_visuals"] == {
        "title_slide_overlay_transparency": 0.9,
        "vocab_slide_overlay_transparency": 0.3,
        "title_card_transparency": 0.1,
        "vocab_card_transparency": 0.8,
    }
    assert captured["build_pptx"]["title_style"]["overlay_transparency"] == 0.9
    assert captured["build_pptx"]["vocab_style"]["show_card"] is True
    assert captured["build_pptx"]["vocab_style"]["card_transparency"] == 0.8
