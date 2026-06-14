from __future__ import annotations

from pathlib import Path


def resolve_csv_path(csv_file: str | Path) -> Path:
    csv_path = Path(csv_file).expanduser().resolve()
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")
    return csv_path


def resolve_output_path(csv_path: Path, output: str | Path | None = None) -> Path:
    if output:
        output_path = Path(output).expanduser().resolve()
    else:
        output_path = csv_path.with_name(csv_path.stem + "_vocab_sets.pptx")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return output_path


def resolve_background_dir(background_dir: str | Path) -> Path:
    bg_dir = Path(background_dir).expanduser().resolve()
    if not bg_dir.exists():
        raise FileNotFoundError(f"Background directory not found: {bg_dir}")
    if not bg_dir.is_dir():
        raise ValueError(f"Background path is not a directory: {bg_dir}")
    return bg_dir
