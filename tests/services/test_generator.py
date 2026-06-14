from pathlib import Path

from mainkata.services import generator


def test_generate_from_inputs_orchestrates_and_returns_outputs(
    monkeypatch, tmp_path: Path
) -> None:
    csv_path = tmp_path / "input.csv"
    output_path = tmp_path / "output.pptx"
    exported_csv_path = tmp_path / "output_selected_terms.csv"

    calls = {}

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
        return {"labels": {"set_prefix": "Set", "vocabulary_suffix": "Vocabulary"}}

    def fake_resolve_title_slide_style(style_config):
        return {
            "overlay_transparency": 0.2,
            "card_transparency": 0.1,
            "show_card": True,
        }

    def fake_resolve_vocab_slide_style(style_config):
        return {
            "overlay_transparency": 0.3,
            "card_transparency": 0.15,
            "show_card": True,
        }

    def fake_validate_visual_options(**kwargs):
        calls["validate_visual_options"] = kwargs

    def fake_read_vocab_csv(path, min_rows):
        calls["read_vocab_csv"] = (path, min_rows)
        return [("inu", "dog"), ("neko", "cat")]

    def fake_random_sets(vocab, set_count, set_size, seed):
        calls["random_sets"] = (vocab, set_count, set_size, seed)
        return [[("inu", "dog")]]

    def fake_build_background_pool(**kwargs):
        calls["build_background_pool"] = kwargs
        return [Path("bg1.jpg")]

    def fake_build_pptx(**kwargs):
        calls["build_pptx"] = kwargs
        return output_path, [(1, "inu", "dog")]

    def fake_write_selected_terms_csv(pptx_output_path, rows):
        calls["write_selected_terms_csv"] = (pptx_output_path, rows)
        return exported_csv_path

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
    monkeypatch.setattr(
        generator, "validate_visual_options", fake_validate_visual_options
    )
    monkeypatch.setattr(generator, "read_vocab_csv", fake_read_vocab_csv)
    monkeypatch.setattr(generator, "random_sets", fake_random_sets)
    monkeypatch.setattr(generator, "build_background_pool", fake_build_background_pool)
    monkeypatch.setattr(generator, "build_pptx", fake_build_pptx)
    monkeypatch.setattr(
        generator, "write_selected_terms_csv", fake_write_selected_terms_csv
    )

    result = generator.generate_from_inputs(
        csv_file="input.csv",
        output="output.pptx",
        style_config_file="style.toml",
        set_count=2,
        set_size=1,
        seed=99,
        primary_side="term",
        show_alternate=True,
        export_selected_terms=True,
        background_dir="backgrounds",
        background_mode="cycle",
    )

    assert result == (output_path, exported_csv_path)
    assert calls["validate_generation_options"] == (2, 1, "term")
    assert calls["resolve_csv_path"] == "input.csv"
    assert calls["load_style_config"] == "style.toml"
    assert calls["read_vocab_csv"] == (csv_path, 1)
    assert calls["write_selected_terms_csv"] == (output_path, [(1, "inu", "dog")])


def test_generate_from_inputs_skips_selected_terms_export_when_disabled(
    monkeypatch, tmp_path: Path
) -> None:
    csv_path = tmp_path / "input.csv"
    output_path = tmp_path / "output.pptx"

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
            "show_card": True,
        },
    )
    monkeypatch.setattr(generator, "validate_visual_options", lambda **k: None)
    monkeypatch.setattr(generator, "read_vocab_csv", lambda *a, **k: [("inu", "dog")])
    monkeypatch.setattr(generator, "random_sets", lambda *a, **k: [[("inu", "dog")]])
    monkeypatch.setattr(generator, "build_background_pool", lambda **k: [])
    monkeypatch.setattr(
        generator, "build_pptx", lambda **k: (output_path, [(1, "inu", "dog")])
    )

    export_called = {"called": False}

    def fake_write_selected_terms_csv(*args, **kwargs):
        export_called["called"] = True
        return tmp_path / "should_not_exist.csv"

    monkeypatch.setattr(
        generator, "write_selected_terms_csv", fake_write_selected_terms_csv
    )

    result = generator.generate_from_inputs(
        csv_file="input.csv",
        export_selected_terms=False,
    )

    assert result == (output_path, None)
    assert export_called["called"] is False
