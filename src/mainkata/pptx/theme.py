from __future__ import annotations

from typing import Any

from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt


def set_shape_fill_transparency(shape, transparency: float) -> None:
    if not 0.0 <= transparency <= 1.0:
        raise ValueError("transparency must be between 0.0 and 1.0")

    fill = shape.fill
    fill.solid()

    solid_fill = shape._element.spPr.solidFill
    if solid_fill is None:
        raise ValueError("Shape does not have a solid fill.")

    color_node = solid_fill.srgbClr
    if color_node is None:
        color_node = solid_fill.schemeClr
    if color_node is None:
        raise ValueError("Solid fill color node not found.")

    for child in list(color_node):
        if child.tag.endswith("alpha"):
            color_node.remove(child)

    alpha = OxmlElement("a:alpha")
    alpha.set("val", str(int((1.0 - transparency) * 100000)))
    color_node.append(alpha)


def apply_font(run, font_spec: dict[str, Any], color: RGBColor) -> None:
    run.font.name = font_spec["name"]
    run.font.size = Pt(font_spec["size"])
    run.font.bold = font_spec["bold"]
    run.font.color.rgb = color


def fit_font_size(text: str) -> int:
    n = len(text)
    if n <= 10:
        return 34
    if n <= 18:
        return 30
    if n <= 26:
        return 26
    if n <= 36:
        return 22
    return 20


def resolve_vocab_primary_font(base_font: dict[str, Any], text: str) -> dict[str, Any]:
    resolved = dict(base_font)
    resolved["size"] = fit_font_size(text)
    return resolved
