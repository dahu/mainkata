from pathlib import Path

import pytest

from mainkata.io.paths import (resolve_background_dir, resolve_csv_path,
                               resolve_output_path)


def test_resolve_csv_path_returns_existing_file(tmp_path: Path) -> None:
    csv_file = tmp_path / "vocab.csv"
    csv_file.write_text(
        "Term,Definition\nCPU,Central Processing Unit\n", encoding="utf-8"
    )

    result = resolve_csv_path(csv_file)

    assert result == csv_file.resolve()


def test_resolve_csv_path_raises_if_file_missing(tmp_path: Path) -> None:
    missing = tmp_path / "missing.csv"

    with pytest.raises(FileNotFoundError, match="CSV file not found:"):
        resolve_csv_path(missing)


def test_resolve_output_path_uses_provided_output(tmp_path: Path) -> None:
    csv_path = (tmp_path / "vocab.csv").resolve()
    output = tmp_path / "custom.pptx"

    result = resolve_output_path(csv_path, output)

    assert result == output.resolve()
    assert result.parent.exists()


def test_resolve_output_path_uses_default_naming_when_output_none(
    tmp_path: Path,
) -> None:
    csv_path = (tmp_path / "lesson.csv").resolve()

    result = resolve_output_path(csv_path, output=None)

    assert result.name == "lesson_vocab_sets.pptx"
    assert result.parent == csv_path.parent
    assert result.parent.exists()


def test_resolve_background_dir_returns_existing_directory(tmp_path: Path) -> None:
    bg_dir = tmp_path / "backgrounds"
    bg_dir.mkdir()

    result = resolve_background_dir(bg_dir)

    assert result == bg_dir.resolve()


def test_resolve_background_dir_raises_if_missing(tmp_path: Path) -> None:
    missing = tmp_path / "does_not_exist"

    with pytest.raises(FileNotFoundError, match="Background directory not found:"):
        resolve_background_dir(missing)


def test_resolve_background_dir_raises_if_not_directory(tmp_path: Path) -> None:
    not_dir = tmp_path / "file.txt"
    not_dir.write_text("not a directory", encoding="utf-8")

    with pytest.raises(ValueError, match="Background directory is not a directory:"):
        resolve_background_dir(not_dir)
