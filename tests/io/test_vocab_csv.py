from pathlib import Path

import pytest

from mainkata.io.vocab_csv import read_vocab_csv


def write_csv(path: Path, content: str) -> Path:
    path.write_text(content, encoding="utf-8")
    return path


def test_read_vocab_csv_reads_valid_rows(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\n" "inu,dog\n" "neko,cat\n",
    )

    result = read_vocab_csv(csv_path, min_rows=2)

    assert result == [
        ("inu", "dog"),
        ("neko", "cat"),
    ]


def test_read_vocab_csv_accepts_case_insensitive_headers(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "term,DEFINITION\n" "aka,red\n" "ao,blue\n",
    )

    result = read_vocab_csv(csv_path, min_rows=2)

    assert result == [
        ("aka", "red"),
        ("ao", "blue"),
    ]


def test_read_vocab_csv_handles_utf8_bom(tmp_path: Path) -> None:
    csv_path = tmp_path / "vocab.csv"
    csv_path.write_text(
        "\ufeffTerm,Definition\ninu,dog\nneko,cat\n",
        encoding="utf-8",
    )

    result = read_vocab_csv(csv_path, min_rows=2)

    assert result == [
        ("inu", "dog"),
        ("neko", "cat"),
    ]


def test_read_vocab_csv_skips_completely_blank_rows(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\n" "inu,dog\n" ",\n" "neko,cat\n",
    )

    result = read_vocab_csv(csv_path, min_rows=2)

    assert result == [
        ("inu", "dog"),
        ("neko", "cat"),
    ]


def test_read_vocab_csv_deduplicates_identical_rows_preserving_order(
    tmp_path: Path,
) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\n"
        "inu,dog\n"
        "neko,cat\n"
        "inu,dog\n"
        "tori,bird\n"
        "neko,cat\n",
    )

    result = read_vocab_csv(csv_path, min_rows=3)

    assert result == [
        ("inu", "dog"),
        ("neko", "cat"),
        ("tori", "bird"),
    ]


def test_read_vocab_csv_strips_whitespace_from_headers_and_values(
    tmp_path: Path,
) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        " Term , Definition \n" " inu , dog \n" " neko, cat\n",
    )

    result = read_vocab_csv(csv_path, min_rows=2)

    assert result == [
        ("inu", "dog"),
        ("neko", "cat"),
    ]


def test_read_vocab_csv_raises_for_missing_required_headers(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Word,Meaning\n" "inu,dog\n",
    )

    with pytest.raises(
        ValueError,
        match="CSV must contain headers for both Term and Definition",
    ):
        read_vocab_csv(csv_path, min_rows=1)


def test_read_vocab_csv_raises_for_incomplete_row_missing_definition(
    tmp_path: Path,
) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\n" "inu,\n",
    )

    with pytest.raises(ValueError, match="Found incomplete row"):
        read_vocab_csv(csv_path, min_rows=1)


def test_read_vocab_csv_raises_for_incomplete_row_missing_term(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\n" ",dog\n",
    )

    with pytest.raises(ValueError, match="Found incomplete row"):
        read_vocab_csv(csv_path, min_rows=1)


def test_read_vocab_csv_raises_when_unique_rows_below_minimum(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\n" "inu,dog\n" "inu,dog\n" "neko,cat\n",
    )

    with pytest.raises(
        ValueError,
        match="CSV must contain at least 3 unique Term/Definition pairs; found 2.",
    ):
        read_vocab_csv(csv_path, min_rows=3)


def test_read_vocab_csv_raises_for_empty_file(tmp_path: Path) -> None:
    csv_path = write_csv(tmp_path / "vocab.csv", "")

    with pytest.raises(
        ValueError,
        match="CSV file is empty or missing a header row.",
    ):
        read_vocab_csv(csv_path, min_rows=1)
