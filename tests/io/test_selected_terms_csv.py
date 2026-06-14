from pathlib import Path

from mainkata.io.selected_terms_csv import write_selected_terms_csv


def test_write_selected_terms_csv_writes_expected_file(tmp_path: Path) -> None:
    output_path = tmp_path / "deck.pptx"
    output_path.write_text("placeholder", encoding="utf-8")

    rows = [
        (1, "apple", "a fruit"),
        (2, "table", "a piece of furniture"),
    ]

    csv_path = write_selected_terms_csv(output_path, rows)

    assert csv_path == tmp_path / "deck_selected_terms.csv"
    assert csv_path.exists()

    content = csv_path.read_text(encoding="utf-8").splitlines()
    assert content == [
        "set_number,term,definition",
        "1,apple,a fruit",
        "2,table,a piece of furniture",
    ]


def test_write_selected_terms_csv_writes_header_only_for_empty_rows(
    tmp_path: Path,
) -> None:
    output_path = tmp_path / "deck.pptx"
    output_path.write_text("placeholder", encoding="utf-8")

    csv_path = write_selected_terms_csv(output_path, [])

    assert csv_path.exists()
    assert csv_path.read_text(encoding="utf-8").splitlines() == [
        "set_number,term,definition",
    ]
