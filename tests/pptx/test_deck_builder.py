from pathlib import Path

from mainkata.pptx import deck_builder


def test_build_source_name_formats_stem() -> None:
    csv_path = Path("/tmp/japanese_n5-lesson_01.csv")
    assert deck_builder._build_source_name(csv_path) == "Japanese N5 Lesson 01"


def test_resolve_slide_text_term_primary_with_alternate() -> None:
    primary, secondary = deck_builder._resolve_slide_text(
        term="inu",
        definition="dog",
        primary_side="term",
        show_alternate=True,
    )
    assert primary == "inu"
    assert secondary == "dog"


def test_resolve_slide_text_definition_primary_without_alternate() -> None:
    primary, secondary = deck_builder._resolve_slide_text(
        term="inu",
        definition="dog",
        primary_side="definition",
        show_alternate=False,
    )
    assert primary == "dog"
    assert secondary is None


def test_build_pptx_returns_selected_rows_and_saves(
    monkeypatch, tmp_path: Path
) -> None:
    calls = {
        "title": [],
        "vocab": [],
        "saved_to": None,
    }

    class FakePresentation:
        def __init__(self) -> None:
            self.slide_width = None
            self.slide_height = None

        def save(self, path: Path) -> None:
            calls["saved_to"] = path

    def fake_add_title_slide(
        prs, set_label, section_title, source_name, labels, style, bg_image
    ):
        calls["title"].append(
            {
                "set_label": set_label,
                "section_title": section_title,
                "source_name": source_name,
                "bg_image": bg_image,
            }
        )

    def fake_add_vocab_slide(prs, primary_text, secondary_text, style, bg_image):
        calls["vocab"].append(
            {
                "primary_text": primary_text,
                "secondary_text": secondary_text,
                "bg_image": bg_image,
            }
        )

    def fake_resolve_background_image(bg_pool, generated_slide_index):
        return bg_pool[generated_slide_index % len(bg_pool)]

    monkeypatch.setattr(deck_builder, "Presentation", FakePresentation)
    monkeypatch.setattr(deck_builder, "add_title_slide", fake_add_title_slide)
    monkeypatch.setattr(deck_builder, "add_vocab_slide", fake_add_vocab_slide)
    monkeypatch.setattr(
        deck_builder, "resolve_background_image", fake_resolve_background_image
    )

    csv_path = tmp_path / "my_vocab.csv"
    output_path = tmp_path / "out.pptx"

    labels = {
        "set_prefix": "Set",
        "vocabulary_suffix": "Vocabulary",
    }
    sets = [
        [("inu", "dog"), ("neko", "cat")],
        [("aka", "red")],
    ]
    title_style = {"dummy": True}
    vocab_style = {"dummy": True}
    bg_pool = [Path("bg1.jpg"), Path("bg2.jpg")]

    returned_output, selected_rows = deck_builder.build_pptx(
        csv_path=csv_path,
        output_path=output_path,
        labels=labels,
        sets=sets,
        primary_side="term",
        show_alternate=True,
        title_style=title_style,
        vocab_style=vocab_style,
        bg_pool=bg_pool,
    )

    assert returned_output == output_path
    assert selected_rows == [
        (1, "inu", "dog"),
        (1, "neko", "cat"),
        (2, "aka", "red"),
    ]
    assert calls["saved_to"] == output_path

    assert len(calls["title"]) == 2
    assert calls["title"][0]["set_label"] == "Set 1"
    assert calls["title"][1]["set_label"] == "Set 2"

    assert len(calls["vocab"]) == 3
    assert calls["vocab"][0]["primary_text"] == "inu"
    assert calls["vocab"][0]["secondary_text"] == "dog"


def test_build_pptx_uses_definition_as_primary_when_requested(
    monkeypatch, tmp_path: Path
) -> None:
    vocab_calls = []

    class FakePresentation:
        def __init__(self) -> None:
            self.slide_width = None
            self.slide_height = None

        def save(self, path: Path) -> None:
            pass

    monkeypatch.setattr(deck_builder, "Presentation", FakePresentation)
    monkeypatch.setattr(deck_builder, "add_title_slide", lambda *a, **k: None)
    monkeypatch.setattr(
        deck_builder,
        "add_vocab_slide",
        lambda prs, primary_text, secondary_text, style, bg_image: vocab_calls.append(
            (primary_text, secondary_text)
        ),
    )
    monkeypatch.setattr(
        deck_builder,
        "resolve_background_image",
        lambda bg_pool, generated_slide_index: None,
    )

    _, selected_rows = deck_builder.build_pptx(
        csv_path=tmp_path / "lesson.csv",
        output_path=tmp_path / "deck.pptx",
        labels={"set_prefix": "Set", "vocabulary_suffix": "Vocabulary"},
        sets=[[("inu", "dog")]],
        primary_side="definition",
        show_alternate=True,
        title_style={},
        vocab_style={},
        bg_pool=[],
    )

    assert vocab_calls == [("dog", "inu")]
    assert selected_rows == [(1, "inu", "dog")]


def test_build_pptx_hides_secondary_text_when_show_alternate_false(
    monkeypatch, tmp_path: Path
) -> None:
    vocab_calls = []

    class FakePresentation:
        def __init__(self) -> None:
            self.slide_width = None
            self.slide_height = None

        def save(self, path: Path) -> None:
            pass

    monkeypatch.setattr(deck_builder, "Presentation", FakePresentation)
    monkeypatch.setattr(deck_builder, "add_title_slide", lambda *a, **k: None)
    monkeypatch.setattr(
        deck_builder,
        "add_vocab_slide",
        lambda prs, primary_text, secondary_text, style, bg_image: vocab_calls.append(
            (primary_text, secondary_text)
        ),
    )
    monkeypatch.setattr(
        deck_builder,
        "resolve_background_image",
        lambda bg_pool, generated_slide_index: None,
    )

    deck_builder.build_pptx(
        csv_path=tmp_path / "lesson.csv",
        output_path=tmp_path / "deck.pptx",
        labels={"set_prefix": "Set", "vocabulary_suffix": "Vocabulary"},
        sets=[[("inu", "dog")]],
        primary_side="term",
        show_alternate=False,
        title_style={},
        vocab_style={},
        bg_pool=[],
    )

    assert vocab_calls == [("inu", None)]
