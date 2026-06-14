from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .types import BackgroundMode, PrimarySide


@dataclass(frozen=True)
class GenerationOptions:
    set_count: int = 6
    set_size: int = 10
    seed: int = 42
    primary_side: PrimarySide = "term"
    show_alternate: bool = True
    export_selected_terms: bool = False


@dataclass(frozen=True)
class BackgroundOptions:
    background_dir: str | Path | None = None
    background_mode: BackgroundMode = "cycle"
    background_image_number: int | None = None
    background_cycle_start: int | None = None
    background_cycle_end: int | None = None


@dataclass(frozen=True)
class VisualOptions:
    title_slide_overlay_transparency: float | None = None
    vocab_slide_overlay_transparency: float | None = None
    show_title_card: bool | None = None
    title_card_transparency: float | None = None
    show_vocab_card: bool | None = None
    vocab_card_transparency: float | None = None
