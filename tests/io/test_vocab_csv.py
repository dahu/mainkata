from pathlib import Path

import pytest

from mainkata.io.vocab_csv import read_vocab_csv


def write_csv(path: Path, content: str) -> Path:
    path.write_text(content, encoding="utf-8")
    return path


def test_read_vocab_csv_reads_valid_rows(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\nCPU,Central Processing Unit\nRAM,Random Access Memory\n",
    )

    rows = read_vocab_csv(csv_path, min_rows=2)

    assert rows == [
        ("CPU", "Central Processing Unit"),
        ("RAM", "Random Access Memory"),
    ]


def test_read_vocab_csv_accepts_case_insensitive_headers(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        " term , definition \nCPU,Central Processing Unit\nRAM,Random Access Memory\n",
    )

    rows = read_vocab_csv(csv_path, min_rows=2)

    assert rows == [
        ("CPU", "Central Processing Unit"),
        ("RAM", "Random Access Memory"),
    ]


def test_read_vocab_csv_skips_fully_blank_rows(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\nCPU,Central Processing Unit\n,\nRAM,Random Access Memory\n",
    )

    rows = read_vocab_csv(csv_path, min_rows=2)

    assert rows == [
        ("CPU", "Central Processing Unit"),
        ("RAM", "Random Access Memory"),
    ]


def test_read_vocab_csv_strips_whitespace_from_values(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "vocab.csv",
        "Term,Definition\n  CPU  ,  Central Processing Unit  \n  RAM , Random Access Memory \n",
    )

    rows = read_vocab_csv(csv_path, min_rows=2)

    assert rows == [
        ("CPU", "Central Processing Unit"),
        ("RAM", "Random Access Memory"),
    ]


def test_read_vocab_csv_rejects_empty_file_or_missing_header_row(
    tmp_path: Path,
) -> None:
    csv_path = write_csv(tmp_path / "empty.csv", "")

    with pytest.raises(ValueError, match="CSV file is empty or missing a header row."):
        read_vocab_csv(csv_path)


def test_read_vocab_csv_rejects_missing_required_headers(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "bad_headers.csv",
        "Word,Meaning\nCPU,Central Processing Unit\n",
    )

    with pytest.raises(
        ValueError,
        match="CSV must contain headers for both Term and Definition \\(case-insensitive\\).",
    ):
        read_vocab_csv(csv_path)


def test_read_vocab_csv_rejects_incomplete_row_missing_term(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "bad_row.csv",
        "Term,Definition\n,Central Processing Unit\n",
    )

    with pytest.raises(
        ValueError,
        match=r"Found incomplete row: Term='', Definition='Central Processing Unit'",
    ):
        read_vocab_csv(csv_path, min_rows=1)


def test_read_vocab_csv_rejects_incomplete_row_missing_definition(
    tmp_path: Path,
) -> None:
    csv_path = write_csv(
        tmp_path / "bad_row.csv",
        "Term,Definition\nCPU,\n",
    )

    with pytest.raises(
        ValueError,
        match=r"Found incomplete row: Term='CPU', Definition=''",
    ):
        read_vocab_csv(csv_path, min_rows=1)


def test_read_vocab_csv_deduplicates_repeated_pairs_preserving_order(
    tmp_path: Path,
) -> None:
    csv_path = write_csv(
        tmp_path / "dupes.csv",
        "Term,Definition\nCPU,Central Processing Unit\nCPU,Central Processing Unit\nRAM,Random Access Memory\nCPU,Central Processing Unit\n",
    )

    rows = read_vocab_csv(csv_path, min_rows=2)

    assert rows == [
        ("CPU", "Central Processing Unit"),
        ("RAM", "Random Access Memory"),
    ]


def test_read_vocab_csv_enforces_minimum_unique_rows(tmp_path: Path) -> None:
    csv_path = write_csv(
        tmp_path / "too_short.csv",
        "Term,Definition\nCPU,Central Processing Unit\nCPU,Central Processing Unit\n",
    )

    with pytest.raises(
        ValueError,
        match="CSV must contain at least 2 unique Term/Definition pairs; found 1.",
    ):
        read_vocab_csv(csv_path, min_rows=2)
