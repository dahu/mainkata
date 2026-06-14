from __future__ import annotations

from .types import BackgroundMode, PrimarySide


def validate_generation_options(
    set_count: int,
    set_size: int,
    primary_side: PrimarySide,
) -> None:
    if set_count < 1:
        raise ValueError("--sets must be at least 1.")
    if set_size < 1:
        raise ValueError("--set-size must be at least 1.")
    if primary_side not in {"term", "definition"}:
        raise ValueError("--primary-side must be either 'term' or 'definition'.")


def validate_background_options(
    background_dir,
    background_mode: BackgroundMode,
    background_image_number: int | None,
    background_cycle_start: int | None,
    background_cycle_end: int | None,
) -> None:
    if background_dir is None:
        return

    if background_mode not in {"fixed", "cycle"}:
        raise ValueError("--background-mode must be 'fixed' or 'cycle'.")

    if background_mode == "fixed":
        if background_image_number is None or background_image_number < 1:
            raise ValueError(
                "--background-image-number must be >= 1 when --background-mode=fixed."
            )
        if background_cycle_start is not None or background_cycle_end is not None:
            raise ValueError(
                "--background-cycle-start and --background-cycle-end cannot be "
                "used with --background-mode=fixed."
            )

    if background_mode == "cycle":
        if background_image_number is not None:
            raise ValueError(
                "--background-image-number cannot be used with "
                "--background-mode=cycle."
            )

        one_range_value = (background_cycle_start is None) != (
            background_cycle_end is None
        )
        if one_range_value:
            raise ValueError(
                "--background-cycle-start and --background-cycle-end must be "
                "provided together."
            )

        if (
            background_cycle_start is not None
            and background_cycle_end is not None
            and background_cycle_start < 1
        ):
            raise ValueError("--background-cycle-start must be at least 1.")

        if (
            background_cycle_start is not None
            and background_cycle_end is not None
            and background_cycle_end < 1
        ):
            raise ValueError("--background-cycle-end must be at least 1.")

        if (
            background_cycle_start is not None
            and background_cycle_end is not None
            and background_cycle_start > background_cycle_end
        ):
            raise ValueError(
                "--background-cycle-start cannot be greater than "
                "--background-cycle-end."
            )


def validate_visual_options(
    title_slide_overlay_transparency: float,
    vocab_slide_overlay_transparency: float,
    title_card_transparency: float,
    vocab_card_transparency: float,
) -> None:
    if not 0.0 <= title_slide_overlay_transparency <= 1.0:
        raise ValueError(
            "--title-slide-overlay-transparency must be between 0.0 and 1.0."
        )
    if not 0.0 <= vocab_slide_overlay_transparency <= 1.0:
        raise ValueError(
            "--vocab-slide-overlay-transparency must be between 0.0 and 1.0."
        )
    if not 0.0 <= title_card_transparency <= 1.0:
        raise ValueError("--title-card-transparency must be between 0.0 and 1.0.")
    if not 0.0 <= vocab_card_transparency <= 1.0:
        raise ValueError("--vocab-card-transparency must be between 0.0 and 1.0.")
