from __future__ import annotations

import sys
from pathlib import Path

import pytest

from mainkata import cli
from mainkata.domain import BackgroundOptions, GenerationOptions, VisualOptions


def write_csv(path: Path, rows: list[tuple[str, str]]) -> None:
    lines = ["Term,Definition"]
    lines.extend(f"{t},{d}" for t, d in rows)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# ----------------------------
# build_parser / basic parsing
# ----------------------------


def test_build_parser_parses_minimal_args(tmp_path: Path) -> None:
    parser = cli.build_parser()
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("CPU", "Processor")] * 10)

    args = parser.parse_args([str(csv_path)])

    assert args.csvfile == str(csv_path)
    assert args.output is None
    assert args.style_config is None
    assert args.sets == 6
    assert args.set_size == 10
    assert args.seed == 42
    assert args.primary_side == "term"
    assert args.hide_alternate is False
    assert args.background_dir is None
    assert args.background_mode == "cycle"
    assert args.background_image_number is None
    assert args.background_cycle_start is None
    assert args.background_cycle_end is None
    assert args.force is False


def test_build_parser_parses_all_groups(tmp_path: Path) -> None:
    parser = cli.build_parser()
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("CPU", "Processor")] * 10)
    out_path = tmp_path / "out.pptx"
    style_path = tmp_path / "style.toml"
    style_path.write_text("", encoding="utf-8")

    argv = [
        str(csv_path),
        "-o",
        str(out_path),
        "--style-config",
        str(style_path),
        "--sets",
        "4",
        "--set-size",
        "8",
        "--seed",
        "99",
        "--export-selected-terms",
        "--primary-side",
        "definition",
        "--hide-alternate",
        "--background-dir",
        str(tmp_path),
        "--background-mode",
        "fixed",
        "--background-image-number",
        "1",
        "--title-slide-overlay-transparency",
        "0.5",
        "--vocab-slide-overlay-transparency",
        "0.3",
        "--hide-title-card",
        "--show-vocab-card",
        "--title-card-transparency",
        "0.4",
        "--vocab-card-transparency",
        "0.2",
        "--force",
    ]

    args = parser.parse_args(argv)

    assert args.csvfile == str(csv_path)
    assert args.output == str(out_path)
    assert args.style_config == str(style_path)
    assert args.sets == 4
    assert args.set_size == 8
    assert args.seed == 99
    assert args.export_selected_terms is True
    assert args.primary_side == "definition"
    assert args.hide_alternate is True
    assert args.background_dir == str(tmp_path)
    assert args.background_mode == "fixed"
    assert args.background_image_number == 1
    assert args.background_cycle_start is None
    assert args.background_cycle_end is None
    assert args.title_slide_overlay_transparency == 0.5
    assert args.vocab_slide_overlay_transparency == 0.3
    assert args.hide_title_card is True
    assert args.show_title_card is False
    assert args.show_vocab_card is True
    assert args.hide_vocab_card is False
    assert args.title_card_transparency == 0.4
    assert args.vocab_card_transparency == 0.2
    assert args.force is True


# ----------------------------
# resolve_optional_bool
# ----------------------------


def test_resolve_optional_bool_returns_true_for_show_flag() -> None:
    parser = cli.build_parser()

    result = cli.resolve_optional_bool(
        parser,
        show_flag=True,
        hide_flag=False,
        show_name="--show-card",
        hide_name="--hide-card",
    )

    assert result is True


def test_resolve_optional_bool_returns_false_for_hide_flag() -> None:
    parser = cli.build_parser()

    result = cli.resolve_optional_bool(
        parser,
        show_flag=False,
        hide_flag=True,
        show_name="--show-card",
        hide_name="--hide-card",
    )

    assert result is False


def test_resolve_optional_bool_returns_none_when_neither_flag() -> None:
    parser = cli.build_parser()

    result = cli.resolve_optional_bool(
        parser,
        show_flag=False,
        hide_flag=False,
        show_name="--show-card",
        hide_name="--hide-card",
    )

    assert result is None


def test_resolve_optional_bool_errors_when_both_flags(
    capsys: pytest.CaptureFixture[str],
) -> None:
    parser = cli.build_parser()

    with pytest.raises(SystemExit) as excinfo:
        cli.resolve_optional_bool(
            parser,
            show_flag=True,
            hide_flag=True,
            show_name="--show-card",
            hide_name="--hide-card",
        )

    assert excinfo.value.code == 2
    captured = capsys.readouterr()
    assert "--show-card and --hide-card cannot be used together." in captured.err


# ----------------------------
# main() happy path
# ----------------------------


def test_main_happy_path_invokes_generator_and_prints_created_files(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("CPU", "Processor")] * 10)

    expected_output_path = cli.resolve_output_path(csv_path, None)
    returned_output_path = tmp_path / "deck.pptx"
    selected_path = tmp_path / "deck_selected.csv"

    captured: dict[str, object] = {}

    def fake_generate_from_inputs(
        csv_file,
        output,
        style_config_file,
        generation: GenerationOptions,
        background: BackgroundOptions,
        visual: VisualOptions,
    ):
        captured["csv_file"] = csv_file
        captured["output"] = output
        captured["style_config_file"] = style_config_file
        captured["generation"] = generation
        captured["background"] = background
        captured["visual"] = visual

        class Result:
            def __init__(self, pptx_path, selected_terms_csv_path):
                self.pptx_path = pptx_path
                self.selected_terms_csv_path = selected_terms_csv_path

        return Result(returned_output_path, selected_path)

    monkeypatch.setattr("mainkata.cli.generate_from_inputs", fake_generate_from_inputs)

    argv = [
        "mainkata",
        str(csv_path),
        "--sets",
        "3",
        "--set-size",
        "10",
        "--seed",
        "7",
        "--export-selected-terms",
        "--primary-side",
        "term",
        "--background-mode",
        "cycle",
        "--force",
    ]
    monkeypatch.setattr(sys, "argv", argv)

    cli.main()

    assert captured["csv_file"] == csv_path
    assert captured["output"] == expected_output_path
    assert captured["style_config_file"] is None

    gen = captured["generation"]
    assert isinstance(gen, GenerationOptions)
    assert gen.set_count == 3
    assert gen.set_size == 10
    assert gen.seed == 7
    assert gen.primary_side == "term"
    assert gen.export_selected_terms is True

    bg = captured["background"]
    assert isinstance(bg, BackgroundOptions)
    assert bg.background_mode == "cycle"

    vis = captured["visual"]
    assert isinstance(vis, VisualOptions)

    captured_out = capsys.readouterr().out
    assert f"Created: {returned_output_path}" in captured_out
    assert f"Created: {selected_path}" in captured_out


# ----------------------------
# main() output path / --force
# ----------------------------


def test_main_errors_when_output_exists_without_force(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("CPU", "Processor")] * 10)
    out_path = tmp_path / "deck.pptx"
    out_path.write_bytes(b"existing")

    monkeypatch.setattr(sys, "argv", ["mainkata", str(csv_path), "-o", str(out_path)])

    with pytest.raises(SystemExit) as excinfo:
        cli.main()

    assert excinfo.value.code == 1
    captured = capsys.readouterr()
    assert "Error: Output file already exists" in captured.err
    assert "Use --force to overwrite." in captured.err


def test_main_allows_overwrite_when_force_set(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("CPU", "Processor")] * 10)
    out_path = tmp_path / "deck.pptx"
    out_path.write_bytes(b"existing")

    def fake_generate_from_inputs(**kwargs):
        class Result:
            def __init__(self, pptx_path):
                self.pptx_path = pptx_path
                self.selected_terms_csv_path = None

        return Result(out_path)

    monkeypatch.setattr("mainkata.cli.generate_from_inputs", fake_generate_from_inputs)
    monkeypatch.setattr(
        sys,
        "argv",
        ["mainkata", str(csv_path), "-o", str(out_path), "--force"],
    )

    cli.main()

    captured = capsys.readouterr().out
    assert f"Created: {out_path}" in captured


# ----------------------------
# main() error handling / debug flag
# ----------------------------


def test_main_exits_with_error_message_on_exception_without_debug(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("CPU", "Processor")] * 10)

    def fake_generate_from_inputs(**kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr("mainkata.cli.generate_from_inputs", fake_generate_from_inputs)
    monkeypatch.setattr(sys, "argv", ["mainkata", str(csv_path)])

    with pytest.raises(SystemExit) as excinfo:
        cli.main()

    assert excinfo.value.code == 1
    captured = capsys.readouterr()
    assert "Error: boom" in captured.err
    # No traceback printed without --debug
    assert "Traceback" not in captured.err


def test_main_with_debug_prints_traceback(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_csv(csv_path, [("CPU", "Processor")] * 10)

    def fake_generate_from_inputs(**kwargs):
        raise RuntimeError("boom")

    monkeypatch.setattr("mainkata.cli.generate_from_inputs", fake_generate_from_inputs)
    monkeypatch.setattr(sys, "argv", ["mainkata", "--debug", str(csv_path)])

    with pytest.raises(SystemExit) as excinfo:
        cli.main()

    assert excinfo.value.code == 1
    captured = capsys.readouterr()
    # Traceback should be printed by traceback.print_exc()
    assert "Traceback" in captured.err
    assert "RuntimeError: boom" in captured.err
