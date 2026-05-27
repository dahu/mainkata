import csv
from pathlib import Path

import pytest
from pptx import Presentation

from src.mainkata.services.generator import (build_pptx, fit_font_size,
                                             generate_from_inputs, random_sets,
                                             read_vocab_csv, resolve_csv_path,
                                             resolve_output_path,
                                             validate_generation_options)


def _write_csv(path: Path, rows, headers=("Term", "Definition")):
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


@pytest.fixture
def vocab_rows():
    return [(f"Term {i}", f"Definition {i}") for i in range(1, 13)]


@pytest.fixture
def vocab_csv(tmp_path: Path, vocab_rows):
    path = tmp_path / "vocab.csv"
    _write_csv(path, vocab_rows)
    return path


class TestResolvePaths:
    def test_resolve_csv_path_returns_resolved_path_for_existing_file(
        self, vocab_csv: Path
    ):
        result = resolve_csv_path(vocab_csv)
        assert result == vocab_csv.resolve()

    def test_resolve_csv_path_raises_for_missing_file(self, tmp_path: Path):
        missing = tmp_path / "missing.csv"
        with pytest.raises(FileNotFoundError, match="CSV file not found"):
            resolve_csv_path(missing)

    def test_resolve_output_path_defaults_next_to_csv(self, vocab_csv: Path):
        result = resolve_output_path(vocab_csv)
        assert result == vocab_csv.with_name("vocab_vocab_sets.pptx")
        assert result.parent.exists()

    def test_resolve_output_path_creates_parent_directories(
        self, vocab_csv: Path, tmp_path: Path
    ):
        output = tmp_path / "nested" / "deeper" / "deck.pptx"
        result = resolve_output_path(vocab_csv, output)
        assert result == output.resolve()
        assert result.parent.exists()


class TestValidateGenerationOptions:
    def test_accepts_valid_options(self):
        validate_generation_options(2, 10, "term")
        validate_generation_options(1, 1, "definition")

    @pytest.mark.parametrize(
        "set_count,set_size,primary_side,error",
        [
            (0, 10, "term", "--sets must be at least 1."),
            (1, 0, "term", "--set-size must be at least 1."),
            (1, 10, "other", "--primary-side must be either 'term' or 'definition'."),
        ],
    )
    def test_rejects_invalid_options(self, set_count, set_size, primary_side, error):
        with pytest.raises(ValueError, match=error):
            validate_generation_options(set_count, set_size, primary_side)


class TestReadVocabCsv:
    def test_reads_case_insensitive_headers(self, tmp_path: Path):
        path = tmp_path / "mixed_headers.csv"
        _write_csv(
            path,
            [("hola", "hello"), ("adios", "goodbye")],
            headers=(" term ", "DEFINITION"),
        )
        result = read_vocab_csv(path, min_rows=2)
        assert result == [("hola", "hello"), ("adios", "goodbye")]

    def test_skips_fully_blank_rows(self, tmp_path: Path):
        path = tmp_path / "blank_rows.csv"
        _write_csv(path, [("a", "1"), ("", ""), ("b", "2")])
        result = read_vocab_csv(path, min_rows=2)
        assert result == [("a", "1"), ("b", "2")]

    def test_deduplicates_rows_preserving_first_occurrence(self, tmp_path: Path):
        path = tmp_path / "dupes.csv"
        _write_csv(path, [("a", "1"), ("b", "2"), ("a", "1"), ("c", "3")])
        result = read_vocab_csv(path, min_rows=3)
        assert result == [("a", "1"), ("b", "2"), ("c", "3")]

    def test_raises_for_missing_header_row(self, tmp_path: Path):
        path = tmp_path / "empty.csv"
        path.write_text("", encoding="utf-8")
        with pytest.raises(
            ValueError, match="CSV file is empty or missing a header row"
        ):
            read_vocab_csv(path)

    def test_raises_for_missing_required_headers(self, tmp_path: Path):
        path = tmp_path / "bad_headers.csv"
        _write_csv(path, [("a", "1")], headers=("Word", "Meaning"))
        with pytest.raises(
            ValueError, match="CSV must contain headers for both Term and Definition"
        ):
            read_vocab_csv(path, min_rows=1)

    def test_raises_for_incomplete_row(self, tmp_path: Path):
        path = tmp_path / "incomplete.csv"
        _write_csv(path, [("a", "1"), ("b", "")])
        with pytest.raises(ValueError, match="Found incomplete row"):
            read_vocab_csv(path, min_rows=1)

    def test_raises_when_unique_row_count_below_minimum(self, tmp_path: Path):
        path = tmp_path / "too_small.csv"
        _write_csv(path, [("a", "1"), ("a", "1"), ("b", "2")])
        with pytest.raises(
            ValueError, match="at least 3 unique Term/Definition pairs; found 2"
        ):
            read_vocab_csv(path, min_rows=3)


class TestRandomSets:
    def test_returns_reproducible_sets_for_same_seed(self, vocab_rows):
        result1 = random_sets(vocab_rows, set_count=3, set_size=4, seed=99)
        result2 = random_sets(vocab_rows, set_count=3, set_size=4, seed=99)
        assert result1 == result2

    def test_returns_requested_number_of_sets_and_items(self, vocab_rows):
        result = random_sets(vocab_rows, set_count=4, set_size=5, seed=7)
        assert len(result) == 4
        assert all(len(s) == 5 for s in result)
        assert all(len(set(s)) == 5 for s in result)

    def test_raises_when_vocab_smaller_than_set_size(self):
        with pytest.raises(ValueError, match="Need at least 5 unique items; got 2"):
            random_sets([("a", "1"), ("b", "2")], set_count=1, set_size=5)


class TestFitFontSize:
    @pytest.mark.parametrize(
        "text,expected",
        [
            ("short", 34),
            ("12345678901", 30),
            ("x" * 19, 26),
            ("x" * 27, 22),
            ("x" * 37, 20),
        ],
    )
    def test_returns_expected_size_by_length_band(self, text, expected):
        assert fit_font_size(text) == expected


class TestBuildPptx:
    def test_build_pptx_creates_expected_slide_count_and_file(
        self, vocab_csv: Path, tmp_path: Path
    ):
        output = tmp_path / "deck.pptx"
        pptx_path, csv_out = build_pptx(
            vocab_csv,
            output,
            set_count=2,
            set_size=3,
            seed=1,
            primary_side="term",
            show_alternate=True,
            export_selected_terms=False,
        )
        assert pptx_path == output
        assert pptx_path.exists()
        assert csv_out is None

        prs = Presentation(pptx_path)
        assert len(prs.slides) == 2 * (1 + 3)

    def test_build_pptx_exports_selected_terms_csv(
        self, vocab_csv: Path, tmp_path: Path
    ):
        output = tmp_path / "deck.pptx"
        pptx_path, csv_out = build_pptx(
            vocab_csv,
            output,
            set_count=2,
            set_size=3,
            seed=2,
            export_selected_terms=True,
        )
        assert pptx_path.exists()
        assert csv_out is not None
        assert csv_out.exists()

        with csv_out.open("r", encoding="utf-8", newline="") as f:
            rows = list(csv.reader(f))

        assert rows[0] == ["set_number", "term", "definition"]
        assert len(rows) == 1 + (2 * 3)
        assert set(row[0] for row in rows[1:]) == {"1", "2"}

    def test_build_pptx_uses_definition_as_primary_when_requested(
        self, vocab_csv: Path, tmp_path: Path
    ):
        output = tmp_path / "deck_definition_first.pptx"
        build_pptx(
            vocab_csv,
            output,
            set_count=1,
            set_size=1,
            seed=0,
            primary_side="definition",
            show_alternate=False,
        )

        prs = Presentation(output)
        vocab_slide = prs.slides[1]
        text_content = "\n".join(
            shape.text
            for shape in vocab_slide.shapes
            if hasattr(shape, "text") and shape.text
        )
        assert "Definition" in text_content
        assert "Term" not in text_content


class TestGenerateFromInputs:
    def test_generate_from_inputs_validates_and_creates_outputs(
        self, vocab_csv: Path, tmp_path: Path
    ):
        output = tmp_path / "generated" / "game_deck.pptx"
        pptx_path, csv_out = generate_from_inputs(
            vocab_csv,
            output=output,
            set_count=1,
            set_size=2,
            seed=5,
            primary_side="term",
            show_alternate=False,
            export_selected_terms=True,
        )
        assert pptx_path == output.resolve()
        assert pptx_path.exists()
        assert csv_out is not None and csv_out.exists()

    def test_generate_from_inputs_raises_before_path_resolution_for_bad_options(
        self, vocab_csv: Path
    ):
        with pytest.raises(ValueError, match="--sets must be at least 1"):
            generate_from_inputs(vocab_csv, set_count=0)
