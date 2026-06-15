import csv
from pathlib import Path

from mainkata.io.selected_terms_csv import write_selected_terms_csv


def test_write_selected_terms_csv_creates_expected_filename(tmp_path: Path) -> None:
    output_path = tmp_path / "lesson_vocab_sets.pptx"

    csv_path = write_selected_terms_csv(
        output_path,
        rows=[(1, "CPU", "Central Processing Unit")],
    )

    assert csv_path == tmp_path / "lesson_vocab_sets_selected_terms.csv"
    assert csv_path.exists()


def test_write_selected_terms_csv_writes_header_and_rows(tmp_path: Path) -> None:
    output_path = tmp_path / "lesson_vocab_sets.pptx"

    csv_path = write_selected_terms_csv(
        output_path,
        rows=[
            (1, "CPU", "Central Processing Unit"),
            (2, "RAM", "Random Access Memory"),
        ],
    )

    with csv_path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        contents = list(reader)

    assert contents == [
        ["Set", "Term", "Definition"],
        ["1", "CPU", "Central Processing Unit"],
        ["2", "RAM", "Random Access Memory"],
    ]


def test_write_selected_terms_csv_overwrites_existing_file(tmp_path: Path) -> None:
    output_path = tmp_path / "lesson_vocab_sets.pptx"
    existing_csv = tmp_path / "lesson_vocab_sets_selected_terms.csv"
    existing_csv.write_text("old,data\n", encoding="utf-8")

    csv_path = write_selected_terms_csv(
        output_path,
        rows=[(3, "GPU", "Graphics Processing Unit")],
    )

    assert csv_path == existing_csv
    assert csv_path.read_text(encoding="utf-8").splitlines() == [
        "Set,Term,Definition",
        "3,GPU,Graphics Processing Unit",
    ]
