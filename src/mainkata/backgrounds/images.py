from __future__ import annotations

from pathlib import Path
from typing import List

from PIL import Image

from mainkata.domain.types import BackgroundMode
from mainkata.io.paths import resolve_background_dir

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg"}


def is_valid_image(path: Path) -> bool:
    try:
        with Image.open(path) as im:
            im.verify()
        return True
    except Exception:
        return False


def list_background_images(background_dir: Path) -> List[Path]:
    candidates = sorted(
        p
        for p in background_dir.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES
    )

    bad: List[Path] = []
    images: List[Path] = []

    for p in candidates:
        if is_valid_image(p):
            images.append(p)
        else:
            bad.append(p)

    if bad:
        bad_list = "\\n".join(str(p) for p in bad)
        raise ValueError(
            "The following files in the background directory are not valid PNG/JPG "
            "images or use unsupported encodings:\\n\\n"
            f"{bad_list}\\n\\n"
            "Please remove or convert them to PNG or JPG."
        )

    if not images:
        raise ValueError(
            f"No valid PNG/JPG images found in background directory: {background_dir}"
        )

    return images


def select_background_pool(
    bg_images: List[Path],
    background_mode: BackgroundMode,
    background_image_number: int | None = None,
    background_cycle_start: int | None = None,
    background_cycle_end: int | None = None,
) -> List[Path]:
    if len(bg_images) == 1:
        return bg_images

    if background_mode == "fixed":
        assert background_image_number is not None
        if background_image_number > len(bg_images):
            raise ValueError(
                f"Requested background image {background_image_number}, "
                f"but only {len(bg_images)} images were found."
            )
        return [bg_images[background_image_number - 1]]

    if background_mode == "cycle":
        if background_cycle_start is None and background_cycle_end is None:
            return bg_images

        assert background_cycle_start is not None
        assert background_cycle_end is not None

        if background_cycle_end > len(bg_images):
            raise ValueError(
                f"Requested background cycle end {background_cycle_end}, "
                f"but only {len(bg_images)} images were found."
            )

        selected = bg_images[background_cycle_start - 1 : background_cycle_end]
        if not selected:
            raise ValueError("Background cycle range did not select any images.")
        return selected

    raise ValueError(f"Unsupported background mode: {background_mode}")


def resolve_background_image(
    bg_pool: List[Path],
    generated_slide_index: int,
) -> Path:
    if not bg_pool:
        raise ValueError("Background image pool is empty.")
    if len(bg_pool) == 1:
        return bg_pool[0]
    return bg_pool[generated_slide_index % len(bg_pool)]


def build_background_pool(
    background_dir: str | Path | None,
    background_mode: BackgroundMode,
    background_image_number: int | None = None,
    background_cycle_start: int | None = None,
    background_cycle_end: int | None = None,
) -> List[Path]:
    if background_dir is None:
        return []

    bg_dir = resolve_background_dir(background_dir)
    bg_images = list_background_images(bg_dir)
    return select_background_pool(
        bg_images,
        background_mode=background_mode,
        background_image_number=background_image_number,
        background_cycle_start=background_cycle_start,
        background_cycle_end=background_cycle_end,
    )
