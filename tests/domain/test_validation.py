import pytest

from mainkata.domain.validation import (validate_background_options,
                                        validate_generation_options,
                                        validate_visual_options)


def test_validate_generation_options_accepts_valid_values() -> None:
    validate_generation_options(
        set_count=6,
        set_size=10,
        primary_side="term",
    )


@pytest.mark.parametrize(
    ("set_count", "set_size", "primary_side", "message"),
    [
        (0, 10, "term", "--sets must be at least 1."),
        (6, 0, "term", "--set-size must be at least 1."),
        (6, 10, "wrong", "--primary-side must be either 'term' or 'definition'."),
    ],
)
def test_validate_generation_options_rejects_invalid_values(
    set_count: int,
    set_size: int,
    primary_side: str,
    message: str,
) -> None:
    with pytest.raises(ValueError, match=message):
        validate_generation_options(
            set_count=set_count,
            set_size=set_size,
            primary_side=primary_side,  # type: ignore[arg-type]
        )


def test_validate_background_options_accepts_none_background_dir() -> None:
    validate_background_options(
        background_dir=None,
        background_mode="cycle",
        background_image_number=None,
        background_cycle_start=None,
        background_cycle_end=None,
    )


def test_validate_background_options_accepts_valid_fixed_mode() -> None:
    validate_background_options(
        background_dir="backgrounds",
        background_mode="fixed",
        background_image_number=1,
        background_cycle_start=None,
        background_cycle_end=None,
    )


def test_validate_background_options_accepts_valid_cycle_mode_without_range() -> None:
    validate_background_options(
        background_dir="backgrounds",
        background_mode="cycle",
        background_image_number=None,
        background_cycle_start=None,
        background_cycle_end=None,
    )


def test_validate_background_options_accepts_valid_cycle_mode_with_range() -> None:
    validate_background_options(
        background_dir="backgrounds",
        background_mode="cycle",
        background_image_number=None,
        background_cycle_start=1,
        background_cycle_end=3,
    )


@pytest.mark.parametrize(
    (
        "background_mode",
        "background_image_number",
        "background_cycle_start",
        "background_cycle_end",
        "message",
    ),
    [
        (
            "fixed",
            None,
            None,
            None,
            "--background-image-number must be >= 1 when --background-mode=fixed.",
        ),
        (
            "fixed",
            0,
            None,
            None,
            "--background-image-number must be >= 1 when --background-mode=fixed.",
        ),
        (
            "fixed",
            1,
            1,
            2,
            "--background-cycle-start and --background-cycle-end cannot be used with --background-mode=fixed.",
        ),
        (
            "cycle",
            1,
            None,
            None,
            "--background-image-number cannot be used with --background-mode=cycle.",
        ),
        (
            "cycle",
            None,
            1,
            None,
            "--background-cycle-start and --background-cycle-end must be provided together.",
        ),
        (
            "cycle",
            None,
            None,
            2,
            "--background-cycle-start and --background-cycle-end must be provided together.",
        ),
        (
            "cycle",
            None,
            0,
            2,
            "--background-cycle-start must be at least 1.",
        ),
        (
            "cycle",
            None,
            1,
            0,
            "--background-cycle-end must be at least 1.",
        ),
        (
            "cycle",
            None,
            3,
            2,
            "--background-cycle-start cannot be greater than --background-cycle-end.",
        ),
    ],
)
def test_validate_background_options_rejects_invalid_combinations(
    background_mode: str,
    background_image_number: int | None,
    background_cycle_start: int | None,
    background_cycle_end: int | None,
    message: str,
) -> None:
    with pytest.raises(ValueError, match=message):
        validate_background_options(
            background_dir="backgrounds",
            background_mode=background_mode,  # type: ignore[arg-type]
            background_image_number=background_image_number,
            background_cycle_start=background_cycle_start,
            background_cycle_end=background_cycle_end,
        )


def test_validate_visual_options_accepts_boundary_values() -> None:
    validate_visual_options(
        title_slide_overlay_transparency=0.0,
        vocab_slide_overlay_transparency=1.0,
        title_card_transparency=0.0,
        vocab_card_transparency=1.0,
    )


@pytest.mark.parametrize(
    ("title_overlay", "vocab_overlay", "title_card", "vocab_card", "message"),
    [
        (
            -0.1,
            0.2,
            0.3,
            0.4,
            "--title-slide-overlay-transparency must be between 0.0 and 1.0.",
        ),
        (
            0.1,
            1.1,
            0.3,
            0.4,
            "--vocab-slide-overlay-transparency must be between 0.0 and 1.0.",
        ),
        (
            0.1,
            0.2,
            -0.1,
            0.4,
            "--title-card-transparency must be between 0.0 and 1.0.",
        ),
        (
            0.1,
            0.2,
            0.3,
            1.1,
            "--vocab-card-transparency must be between 0.0 and 1.0.",
        ),
    ],
)
def test_validate_visual_options_rejects_out_of_range_values(
    title_overlay: float,
    vocab_overlay: float,
    title_card: float,
    vocab_card: float,
    message: str,
) -> None:
    with pytest.raises(ValueError, match=message):
        validate_visual_options(
            title_slide_overlay_transparency=title_overlay,
            vocab_slide_overlay_transparency=vocab_overlay,
            title_card_transparency=title_card,
            vocab_card_transparency=vocab_card,
        )
