"""RGB / color helpers for PDF → Excel."""

from __future__ import annotations

from typing import Any, Iterable, Tuple


def int_to_rgb(value: int) -> Tuple[int, int, int]:
    """Convert PDF color int (0xRRGGBB) to 0–255 RGB."""
    if value < 0:
        value = 0
    r = (value >> 16) & 0xFF
    g = (value >> 8) & 0xFF
    b = value & 0xFF
    return (r, g, b)


def rgb_to_hex(rgb: Iterable[int]) -> str:
    r, g, b = _norm_rgb(rgb)
    return f"{r:02X}{g:02X}{b:02X}"


def _norm_rgb(rgb: Iterable[int]) -> Tuple[int, int, int]:
    t = tuple(rgb)
    if len(t) != 3:
        return (0, 0, 0)
    return tuple(max(0, min(255, int(x))) for x in t)  # type: ignore[return-value]


def fitz_color_to_rgb(color: Any) -> Tuple[int, int, int]:
    """Normalize PyMuPDF color (int, tuple, or list) to RGB."""
    if color is None:
        return (0, 0, 0)
    if isinstance(color, (int, float)):
        return int_to_rgb(int(color))
    if isinstance(color, (list, tuple)) and len(color) >= 3:
        return (int(color[0] * 255), int(color[1] * 255), int(color[2] * 255))
    return (0, 0, 0)
