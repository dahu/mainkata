import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

from src.mainkata import cli  # adjust if your package name/module path differs


def _run_main_with_args(monkeypatch, args):
    """Helper to run cli.main() with a fake argv."""
    monkeypatch.setattr(sys, "argv", ["mainkata"] + args)
    cli.main()


class TestCliHappyPath:
    def test_calls_generate_with_expected_arguments(
        self, monkeypatch, tmp_path, capsys
    ):
        # Arrange: fake CSV and output paths
        csv_path = tmp_path / "input.csv"
        csv_path.write_text("Term,Definition\nfoo,bar\nfoo2,bar2\n", encoding="utf-8")
        output_path = tmp_path / "deck.pptx"

        called = {}

        def fake_generate_from_inputs(**kwargs):
            called.update(kwargs)
            pptx_path = output_path
            csv_out = None
            return pptx_path, csv_out

        # Patch the imported generate function inside cli.py
        monkeypatch.setattr(cli, "generate_from_inputs", fake_generate_from_inputs)

        # Act
        _run_main_with_args(
            monkeypatch,
            [
                str(csv_path),
                "-o",
                str(output_path),
                "--sets",
                "3",
                "--set-size",
                "5",
                "--seed",
                "99",
                "--primary-side",
                "definition",
                "--export-selected-terms",
                "--hide-alternate",
            ],
        )

        # Assert: function was called with mapped arguments
        assert called["csv_file"] == str(csv_path)
        assert called["output"] == str(output_path)
        assert called["set_count"] == 3
        assert called["set_size"] == 5
        assert called["seed"] == 99
        assert called["primary_side"] == "definition"
        assert called["show_alternate"] is False
        assert called["export_selected_terms"] is True

        # Assert: output printed
        out = capsys.readouterr().out
        assert f"Created: {output_path}" in out


class TestCliOverwriteHandling:
    def test_refuses_to_overwrite_existing_output_without_force(
        self, monkeypatch, tmp_path, capsys
    ):
        csv_path = tmp_path / "input.csv"
        csv_path.write_text("Term,Definition\nfoo,bar\n", encoding="utf-8")

        out_path = tmp_path / "existing.pptx"
        out_path.write_bytes(b"stub")

        monkeypatch.setattr(
            cli,
            "generate_from_inputs",
            lambda **_: (_ for _ in ()).throw(
                AssertionError("generate_from_inputs should not be called")
            ),
        )

        monkeypatch.setattr(
            sys,
            "argv",
            [
                "mainkata",
                str(csv_path),
                "--output",
                str(out_path),
            ],
        )

        with pytest.raises(SystemExit) as e:
            cli.main()

        assert e.value.code == 1

        captured = capsys.readouterr()
        combined = captured.out + captured.err
        assert "Output file already exists" in combined
        assert "Use --force to overwrite" in combined


class TestCliErrorHandling:
    def test_exits_with_error_when_generate_raises(self, monkeypatch, tmp_path, capsys):
        csv_path = tmp_path / "input.csv"
        csv_path.write_text("Term,Definition\nfoo,bar\n", encoding="utf-8")

        def fake_generate_from_inputs(**kwargs):
            raise ValueError("bad CSV")

        monkeypatch.setattr(cli, "generate_from_inputs", fake_generate_from_inputs)
        monkeypatch.setattr(sys, "argv", ["mainkata", str(csv_path)])

        with pytest.raises(SystemExit) as e:
            cli.main()

        assert e.value.code == 1

        captured = capsys.readouterr()
        combined = captured.out + captured.err
        assert "Error: bad CSV" in combined
