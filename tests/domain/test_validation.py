import pytest

from mainkata.domain.validation import (
    validate_background_options,
    validate_generation_options,
    validate_visual_options,
)


@pytest.mark.parametrize(
    ("set_count, set_size, primary_side, message"),
    [
        (0, 10, "term", "--sets must be at least 1."),
        (2, 0, "term", "--set-size must be at least 1."),
        (2, 10, "front", "--primary-side must be either 'term' or 'definition'."),
    ],
)
def test_validate_generation_options_rejects_invalid_values(
    set_count,
    set_size,
    primary_side,
    message,
) -> None:
    with pytest.raises(ValueError, match=message):
        validate_generation_options(set_count, set_size, primary_side)


def test_validate_generation_options_accepts_valid_values() -> None:
    validate_generation_options(2, 10, "term")
    validate_generation_options(1, 1, "definition")


@pytest.mark.parametrize(
    ("kwargs, message"),
    [
        (
            {
                "background_dir": "bg",
                "background_mode": "fixed",
                "background_image_number": None,
                "background_cycle_start": None,
                "background_cycle_end": None,
            },
            "--background-image-number must be >= 1 when ",
        ),
        (
            {
                "background_dir": "bg",
                "background_mode": "fixed",
                "background_image_number": 1,
                "background_cycle_start": 1,
                "background_cycle_end": 3,
            },
            "--background-cycle-start and --background-cycle-end cannot be ",
        ),
        (
            {
                "background_dir": "bg",
                "background_mode": "cycle",
                "background_image_number": 2,
                "background_cycle_start": None,
                "background_cycle_end": None,
            },
            "--background-image-number cannot be used with ",
        ),
        (
            {
                "background_dir": "bg",
                "background_mode": "cycle",
                "background_image_number": None,
                "background_cycle_start": 1,
                "background_cycle_end": None,
            },
            "--background-cycle-start and --background-cycle-end must be ",
        ),
        (
            {
                "background_dir": "bg",
                "background_mode": "cycle",
                "background_image_number": None,
                "background_cycle_start": 3,
                "background_cycle_end": 2,
            },
            "--background-cycle-start cannot be greater than ",
        ),
    ],
)
def test_validate_background_options_rejects_invalid_combinations(
    kwargs, message
) -> None:
    with pytest.raises(ValueError, match=message):
        validate_background_options(**kwargs)


def test_validate_background_options_allows_none_background_dir() -> None:
    validate_background_options(
        background_dir=None,
        background_mode="cycle",
        background_image_number=None,
        background_cycle_start=None,
        background_cycle_end=None,
    )


def test_validate_background_options_accepts_valid_fixed_mode() -> None:
    validate_background_options(
        background_dir="bg",
        background_mode="fixed",
        background_image_number=1,
        background_cycle_start=None,
        background_cycle_end=None,
    )


def test_validate_background_options_accepts_valid_cycle_mode() -> None:
    validate_background_options(
        background_dir="bg",
        background_mode="cycle",
        background_image_number=None,
        background_cycle_start=1,
        background_cycle_end=3,
    )


@pytest.mark.parametrize(
    ("kwargs, message"),
    [
        (
            {
                "title_slide_overlay_transparency": -0.1,
                "vocab_slide_overlay_transparency": 0.2,
                "title_card_transparency": 0.3,
                "vocab_card_transparency": 0.4,
            },
            "--title-slide-overlay-transparency must be between 0.0 and 1.0.",
        ),
        (
            {
                "title_slide_overlay_transparency": 0.1,
                "vocab_slide_overlay_transparency": 1.2,
                "title_card_transparency": 0.3,
                "vocab_card_transparency": 0.4,
            },
            "--vocab-slide-overlay-transparency must be between 0.0 and 1.0.",
        ),
        (
            {
                "title_slide_overlay_transparency": 0.1,
                "vocab_slide_overlay_transparency": 0.2,
                "title_card_transparency": -0.01,
                "vocab_card_transparency": 0.4,
            },
            "--title-card-transparency must be between 0.0 and 1.0.",
        ),
        (
            {
                "title_slide_overlay_transparency": 0.1,
                "vocab_slide_overlay_transparency": 0.2,
                "title_card_transparency": 0.3,
                "vocab_card_transparency": 1.01,
            },
            "--vocab-card-transparency must be between 0.0 and 1.0.",
        ),
    ],
)
def test_validate_visual_options_rejects_out_of_range_values(kwargs, message) -> None:
    with pytest.raises(ValueError, match=message):
        validate_visual_options(**kwargs)


def test_validate_visual_options_accepts_boundary_values() -> None:
    validate_visual_options(
        title_slide_overlay_transparency=0.0,
        vocab_slide_overlay_transparency=1.0,
        title_card_transparency=0.0,
        vocab_card_transparency=1.0,
    )
