from __future__ import annotations

from pathlib import Path

import pytest

from mainkata.domain.options import BackgroundOptions, GenerationOptions
from mainkata.io.paths import (resolve_background_dir, resolve_csv_path,
                               resolve_output_path)
from mainkata.io.selected_terms_csv import write_selected_terms_csv
from mainkata.io.vocab_csv import read_vocab_csv
from mainkata.services.generator import generate_from_inputs


def write_csv_raw(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


def write_vocab(path: Path, rows: list[tuple[str, str]]) -> None:
    lines = ["Term,Definition"]
    lines.extend(f"{t},{d}" for t, d in rows)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# ----------------------------
# read_vocab_csv
# ----------------------------


def test_read_vocab_csv_happy_path_min_rows(tmp_path: Path) -> None:
    csv_path = tmp_path / "vocab.csv"
    rows = [(f"T{i}", f"D{i}") for i in range(1, 11)]
    write_vocab(csv_path, rows)

    result = read_vocab_csv(csv_path, min_rows=10)

    assert result == rows


def test_read_vocab_csv_handles_utf8_bom_and_header_casing(tmp_path: Path) -> None:
    csv_path = tmp_path / "vocab.csv"
    # utf-8-sig BOM + odd casing and whitespace
    content = "\ufeff term , DEFINITION \nCPU,Processor\nRam,Memory\n"
    write_csv_raw(csv_path, content)

    result = read_vocab_csv(csv_path, min_rows=2)

    assert result == [("CPU", "Processor"), ("Ram", "Memory")]


def test_read_vocab_csv_raises_on_empty_file(tmp_path: Path) -> None:
    csv_path = tmp_path / "empty.csv"
    write_csv_raw(csv_path, "")

    with pytest.raises(
        ValueError, match=r"CSV file is empty or missing a header row\."
    ):
        read_vocab_csv(csv_path)


def test_read_vocab_csv_requires_term_and_definition_headers(tmp_path: Path) -> None:
    csv_path = tmp_path / "bad.csv"
    write_csv_raw(csv_path, "Word,Meaning\nCPU,Processor\n")

    with pytest.raises(
        ValueError,
        match=r"CSV must contain headers for both Term and Definition",
    ):
        read_vocab_csv(csv_path)


def test_read_vocab_csv_skips_blank_rows(tmp_path: Path) -> None:
    csv_path = tmp_path / "vocab.csv"
    content = "Term,Definition\nCPU,Processor\n,,\nRAM,Memory\n"
    write_csv_raw(csv_path, content)

    result = read_vocab_csv(csv_path, min_rows=2)

    assert result == [("CPU", "Processor"), ("RAM", "Memory")]


def test_read_vocab_csv_rejects_incomplete_rows(tmp_path: Path) -> None:
    csv_path = tmp_path / "vocab.csv"
    content = "Term,Definition\nCPU,\n,RAM\nOK,Value\n"
    write_csv_raw(csv_path, content)

    with pytest.raises(
        ValueError,
        match=r"Found incomplete row: Term=",
    ):
        read_vocab_csv(csv_path, min_rows=1)


def test_read_vocab_csv_deduplicates_term_definition_pairs(tmp_path: Path) -> None:
    csv_path = tmp_path / "vocab.csv"
    content = "Term,Definition\nCPU,Processor\nCPU,Processor\nCPU,Central Unit\n"
    write_csv_raw(csv_path, content)

    result = read_vocab_csv(csv_path, min_rows=2)

    assert result == [("CPU", "Processor"), ("CPU", "Central Unit")]


def test_read_vocab_csv_enforces_min_unique_rows(tmp_path: Path) -> None:
    csv_path = tmp_path / "vocab.csv"
    content = "Term,Definition\nCPU,Processor\nCPU,Processor\n"
    write_csv_raw(csv_path, content)

    with pytest.raises(
        ValueError,
        match=r"CSV must contain at least 3 unique Term/Definition pairs; found 1\.",
    ):
        read_vocab_csv(csv_path, min_rows=3)


# ----------------------------
# write_selected_terms_csv
# ----------------------------


def test_write_selected_terms_csv_creates_companion_file(tmp_path: Path) -> None:
    pptx_path = tmp_path / "deck.pptx"
    pptx_path.write_bytes(b"dummy")

    rows = [
        (1, "CPU", "Processor"),
        (1, "RAM", "Memory"),
        (2, "Disk", "Storage"),
    ]

    csv_out = write_selected_terms_csv(pptx_path, rows)

    assert csv_out.parent == pptx_path.parent
    assert csv_out.suffix == ".csv"
    assert csv_out.exists()

    text = csv_out.read_text(encoding="utf-8").strip().splitlines()
    assert text[0] == "Set,Term,Definition"
    assert text[1:] == [
        "1,CPU,Processor",
        "1,RAM,Memory",
        "2,Disk,Storage",
    ]


# ----------------------------
# paths: resolve_csv_path, resolve_output_path, resolve_background_dir
# ----------------------------


def test_resolve_csv_path_expands_user_and_requires_existing_file(
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "vocab.csv"
    write_vocab(csv_path, [("CPU", "Processor")] * 10)

    resolved = resolve_csv_path(csv_path)

    assert resolved == csv_path.resolve()

    with pytest.raises(FileNotFoundError, match=r"CSV file not found:"):
        resolve_csv_path(tmp_path / "missing.csv")


def test_resolve_output_path_derives_default_name_from_csv(tmp_path: Path) -> None:
    csv_path = tmp_path / "my_terms.csv"
    csv_path.write_text("Term,Definition\nCPU,Processor\n", encoding="utf-8")

    output = resolve_output_path(csv_path, output=None)

    assert output.parent == csv_path.parent
    assert output.name == "my_terms_vocab_sets.pptx"


def test_resolve_output_path_uses_explicit_output_and_expands_user(
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "terms.csv"
    csv_path.write_text("Term,Definition\nCPU,Processor\n", encoding="utf-8")

    explicit = tmp_path / "custom.pptx"
    output = resolve_output_path(csv_path, output=explicit)

    assert output == explicit.resolve()


def test_resolve_background_dir_expands_user_and_requires_directory(
    tmp_path: Path,
) -> None:
    bg_dir = tmp_path / "backgrounds"
    bg_dir.mkdir()

    resolved = resolve_background_dir(bg_dir)

    assert resolved == bg_dir.resolve()

    file_path = tmp_path / "file.txt"
    file_path.write_text("x", encoding="utf-8")

    with pytest.raises(ValueError, match=r"Background directory is not a directory:"):
        resolve_background_dir(file_path)

    with pytest.raises(FileNotFoundError, match=r"Background directory not found:"):
        resolve_background_dir(tmp_path / "missing_dir")


# ----------------------------
# light integration: generate_from_inputs
# ----------------------------


def test_generate_from_inputs_creates_pptx_and_selected_terms_csv(
    tmp_path: Path,
) -> None:
    csv_path = tmp_path / "vocab.csv"
    # Need at least set_size unique rows
    vocab_rows = [(f"T{i}", f"D{i}") for i in range(1, 7)]
    write_vocab(csv_path, vocab_rows)

    output = tmp_path / "deck.pptx"

    generation = GenerationOptions(
        set_count=2,
        set_size=3,
        seed=123,
        primary_side="term",
        show_alternate=True,
        export_selected_terms=True,
    )

    background = BackgroundOptions(background_dir=None)

    result = generate_from_inputs(
        csv_file=csv_path,
        output=output,
        generation=generation,
        background=background,
    )

    assert result.pptx_path == output
    assert output.exists()

    selected_csv = result.selected_terms_csv_path
    assert selected_csv is not None
    assert selected_csv.exists()

    # Basic sanity: slide count matches (set_count * (1 title + set_size))
    from pptx import Presentation

    prs = Presentation(output)
    expected_slides = generation.set_count * (1 + generation.set_size)
    assert len(prs.slides) == expected_slides
