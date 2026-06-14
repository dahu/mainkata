from __future__ import annotations

import os
from copy import deepcopy
from pathlib import Path
from typing import Any, TypedDict

from pptx.dml.color import RGBColor

try:
    import tomllib
except ModuleNotFoundError:
    import tomli as tomllib


class FontSpec(TypedDict):
    name: str
    size: int
    bold: bool


class ColorPalette(TypedDict):
    bg: RGBColor
    blue: RGBColor
    text: RGBColor
    subtext: RGBColor
    coral: RGBColor
    white: RGBColor


class TitleSlideStyle(TypedDict):
    colors: ColorPalette
    pill_font: FontSpec
    main_font: FontSpec
    section_font: FontSpec
    body_font: FontSpec
    overlay_transparency: float
    show_card: bool
    card_transparency: float


class VocabSlideStyle(TypedDict):
    colors: ColorPalette
    primary_font: FontSpec
    secondary_font: FontSpec
    overlay_transparency: float
    show_card: bool
    card_transparency: float


DEFAULT_STYLE_CONFIG: dict[str, Any] = {
    "labels": {
        "game_title": "Vocabulary Games",
        "source_prefix": "Source",
        "set_prefix": "Set",
        "vocabulary_suffix": "Vocabulary",
    },
    "palettes": {
        "colors": {
            "default": {
                "bg": "F0F7FF",
                "blue": "2563EB",
                "text": "0F172A",
                "subtext": "475569",
                "coral": "F97066",
                "white": "FFFFFF",
            }
        },
        "fonts": {
            "title_pill": {"name": "Aptos", "size": 16, "bold": True},
            "title_main": {"name": "Aptos Display", "size": 28, "bold": True},
            "title_section": {"name": "Aptos Display", "size": 24, "bold": True},
            "body": {"name": "Aptos", "size": 16, "bold": False},
            "vocab_primary": {"name": "Aptos Display", "size": 24, "bold": True},
            "vocab_secondary": {"name": "Aptos", "size": 20, "bold": False},
        },
    },
    "styles": {
        "title_slide": {
            "color_palette": "default",
            "pill_font": "title_pill",
            "main_font": "title_main",
            "section_font": "title_section",
            "body_font": "body",
            "overlay_transparency": 0.22,
            "show_card": True,
            "card_transparency": 0.18,
        },
        "vocab_slide": {
            "color_palette": "default",
            "primary_font": "vocab_primary",
            "secondary_font": "vocab_secondary",
            "overlay_transparency": 0.22,
            "show_card": True,
            "card_transparency": 0.18,
        },
    },
}


def deep_merge_dicts(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged = deepcopy(base)
    for key, value in override.items():
        if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
            merged[key] = deep_merge_dicts(merged[key], value)
        else:
            merged[key] = value
    return merged


def get_default_style_config_path() -> Path:
    xdg_config_home = os.environ.get("XDG_CONFIG_HOME", "").strip()
    config_home = (
        Path(xdg_config_home).expanduser()
        if xdg_config_home
        else Path.home() / ".config"
    )
    return config_home / "mainkata" / "style.toml"


def resolve_style_config_path(
    style_config_file: str | Path | None = None,
) -> Path | None:
    if style_config_file is not None:
        config_path = Path(style_config_file).expanduser().resolve()
        if not config_path.exists():
            raise FileNotFoundError(f"Style config file not found: {config_path}")
        if not config_path.is_file():
            raise ValueError(f"Style config path is not a file: {config_path}")
        return config_path

    default_path = get_default_style_config_path().resolve()
    if default_path.exists() and default_path.is_file():
        return default_path
    return None


def load_style_config(style_config_file: str | Path | None = None) -> dict[str, Any]:
    config_path = resolve_style_config_path(style_config_file)
    if config_path is None:
        return deepcopy(DEFAULT_STYLE_CONFIG)

    with config_path.open("rb") as f:
        loaded = tomllib.load(f)

    return deep_merge_dicts(DEFAULT_STYLE_CONFIG, loaded)


def hex_to_rgb_color(value: str) -> RGBColor:
    text = value.strip().lstrip("#")
    if len(text) != 6:
        raise ValueError(f"Invalid hex color value: {value!r}")
    try:
        return RGBColor.from_string(text.upper())
    except ValueError as exc:
        raise ValueError(f"Invalid hex color value: {value!r}") from exc


def resolve_color_palette(
    style_config: dict[str, Any], palette_name: str
) -> ColorPalette:
    palettes = style_config["palettes"]["colors"]
    if palette_name not in palettes:
        raise ValueError(f"Unknown color palette: {palette_name}")
    return {
        key: hex_to_rgb_color(value) for key, value in palettes[palette_name].items()
    }


def resolve_font_palette(style_config: dict[str, Any], font_name: str) -> FontSpec:
    palettes = style_config["palettes"]["fonts"]
    if font_name not in palettes:
        raise ValueError(f"Unknown font palette: {font_name}")
    font = palettes[font_name]
    return {
        "name": str(font["name"]),
        "size": int(font["size"]),
        "bold": bool(font["bold"]),
    }


def resolve_title_slide_style(style_config: dict[str, Any]) -> TitleSlideStyle:
    style = style_config["styles"]["title_slide"]
    return {
        "colors": resolve_color_palette(style_config, style["color_palette"]),
        "pill_font": resolve_font_palette(style_config, style["pill_font"]),
        "main_font": resolve_font_palette(style_config, style["main_font"]),
        "section_font": resolve_font_palette(style_config, style["section_font"]),
        "body_font": resolve_font_palette(style_config, style["body_font"]),
        "overlay_transparency": float(style["overlay_transparency"]),
        "show_card": bool(style["show_card"]),
        "card_transparency": float(style["card_transparency"]),
    }


def resolve_vocab_slide_style(style_config: dict[str, Any]) -> VocabSlideStyle:
    style = style_config["styles"]["vocab_slide"]
    return {
        "colors": resolve_color_palette(style_config, style["color_palette"]),
        "primary_font": resolve_font_palette(style_config, style["primary_font"]),
        "secondary_font": resolve_font_palette(style_config, style["secondary_font"]),
        "overlay_transparency": float(style["overlay_transparency"]),
        "show_card": bool(style["show_card"]),
        "card_transparency": float(style["card_transparency"]),
    }
