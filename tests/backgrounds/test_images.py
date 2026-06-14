from pathlib import Path

import pytest
from PIL import Image

from mainkata.backgrounds import (build_background_pool,
                                  list_background_images,
                                  resolve_background_image,
                                  select_background_pool)
from mainkata.domain import BackgroundOptions


def create_image(path: Path, size: tuple[int, int] = (10, 10)) -> None:
    img = Image.new("RGB", size, color=(255, 0, 0))
    img.save(path)


def test_list_background_images_returns_valid_images(tmp_path: Path) -> None:
    valid_png = tmp_path / "bg1.png"
    valid_jpg = tmp_path / "bg2.jpg"
    invalid_txt = tmp_path / "not_image.txt"

    create_image(valid_png)
    create_image(valid_jpg)
    invalid_txt.write_text("not an image", encoding="utf-8")

    images = list_background_images(tmp_path)

    # Only valid images, sorted
    assert images == [valid_jpg.resolve(), valid_png.resolve()] or images == [
        valid_png.resolve(),
        valid_jpg.resolve(),
    ]


def test_list_background_images_raises_if_invalid_files_present(tmp_path: Path) -> None:
    valid_png = tmp_path / "bg1.png"
    bad_png = tmp_path / "bg_bad.png"

    create_image(valid_png)
    bad_png.write_text("corrupted image", encoding="utf-8")

    with pytest.raises(ValueError) as excinfo:
        list_background_images(tmp_path)

    message = str(excinfo.value)
    assert "not valid PNG/JPG images or use unsupported encodings" in message
    # Should list the bad file path somewhere in the error message
    assert "bg_bad.png" in message


def test_list_background_images_raises_if_no_images_found(tmp_path: Path) -> None:
    (tmp_path / "notes.txt").write_text("no images here", encoding="utf-8")

    with pytest.raises(ValueError) as excinfo:
        list_background_images(tmp_path)

    assert "No valid PNG/JPG images found in background directory" in str(excinfo.value)


def test_select_background_pool_fixed_mode_selects_single_image() -> None:
    images = [Path(f"bg{i}.png") for i in range(1, 4)]

    pool = select_background_pool(
        bg_images=images,
        background_mode="fixed",
        background_image_number=2,
        background_cycle_start=None,
        background_cycle_end=None,
    )

    assert pool == [Path("bg2.png")]


def test_select_background_pool_fixed_mode_raises_if_index_too_large() -> None:
    images = [Path(f"bg{i}.png") for i in range(1, 3)]

    with pytest.raises(
        ValueError, match="Requested background image 3, but only 2 images were found."
    ):
        select_background_pool(
            bg_images=images,
            background_mode="fixed",
            background_image_number=3,
            background_cycle_start=None,
            background_cycle_end=None,
        )


def test_select_background_pool_cycle_mode_full_range_when_no_bounds() -> None:
    images = [Path(f"bg{i}.png") for i in range(1, 4)]

    pool = select_background_pool(
        bg_images=images,
        background_mode="cycle",
        background_image_number=None,
        background_cycle_start=None,
        background_cycle_end=None,
    )

    assert pool == images


def test_select_background_pool_cycle_mode_with_range() -> None:
    images = [Path(f"bg{i}.png") for i in range(1, 6)]

    pool = select_background_pool(
        bg_images=images,
        background_mode="cycle",
        background_image_number=None,
        background_cycle_start=2,
        background_cycle_end=4,
    )

    assert pool == images[1:4]


def test_select_background_pool_cycle_mode_raises_if_end_too_large() -> None:
    images = [Path(f"bg{i}.png") for i in range(1, 3)]

    with pytest.raises(
        ValueError,
        match="Requested background cycle end 3, but only 2 images were found.",
    ):
        select_background_pool(
            bg_images=images,
            background_mode="cycle",
            background_image_number=None,
            background_cycle_start=1,
            background_cycle_end=3,
        )


def test_select_background_pool_cycle_mode_raises_if_range_empty() -> None:
    images = [Path(f"bg{i}.png") for i in range(1, 3)]

    with pytest.raises(
        ValueError, match="Background cycle range did not select any images."
    ):
        select_background_pool(
            bg_images=images,
            background_mode="cycle",
            background_image_number=None,
            background_cycle_start=3,
            background_cycle_end=2,
        )


def test_resolve_background_image_raises_if_pool_empty() -> None:
    with pytest.raises(ValueError, match="Background image pool is empty."):
        resolve_background_image([], generated_slide_index=0)


def test_resolve_background_image_returns_single_when_only_one() -> None:
    image = Path("bg1.png")

    result = resolve_background_image([image], generated_slide_index=5)

    assert result == image


def test_resolve_background_image_cycles_through_pool() -> None:
    images = [Path(f"bg{i}.png") for i in range(1, 4)]

    # Index modulo length should map indexes to images
    assert resolve_background_image(images, 0) == images[0]
    assert resolve_background_image(images, 1) == images[1]
    assert resolve_background_image(images, 2) == images[2]
    assert resolve_background_image(images, 3) == images[0]
    assert resolve_background_image(images, 4) == images[1]


def test_build_background_pool_returns_empty_list_when_no_dir() -> None:
    options = BackgroundOptions(
        background_dir=None,
        background_mode="cycle",
    )

    pool = build_background_pool(options)

    assert pool == []


def test_build_background_pool_uses_resolved_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    bg_dir = tmp_path / "backgrounds"
    bg_dir.mkdir()
    img1 = bg_dir / "bg1.png"
    create_image(img1)

    # Ensure resolve_background_dir points to our temp dir
    from mainkata.io.paths import \
        resolve_background_dir as real_resolve_background_dir

    def fake_resolve_background_dir(path: str | Path) -> Path:
        # ignore the passed path, always return our temp dir
        return real_resolve_background_dir(bg_dir)

    monkeypatch.setattr(
        "mainkata.backgrounds.images.resolve_background_dir",
        fake_resolve_background_dir,
    )

    options = BackgroundOptions(
        background_dir="ignored",
        background_mode="cycle",
    )

    pool = build_background_pool(options)

    assert pool == [img1.resolve()]
