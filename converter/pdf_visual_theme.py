"""Derive Excel styling (fonts, fills, row scale, print hints) from PDF TextBlock stats."""

from __future__ import annotations

from dataclasses import dataclass
from statistics import median
from typing import TYPE_CHECKING, Any, List, Optional, Tuple

from .color_utils import rgb_to_hex

if TYPE_CHECKING:
    from .pdf_parser import ParsedPDF, TextBlock

# Match previous hardcoded TRACES palette in excel_builder
_DEFAULT_PRIMARY = "1D4E89"
_DEFAULT_SUB = "D9E1F2"
_DEFAULT_MASTER = "E8EDF5"
_DEFAULT_TITLE = "004B87"
_DEFAULT_ZEBRA = "F7F9FC"
_DEFAULT_BODY_FONT = "000000"
_DEFAULT_HEADER_ON_PRIMARY = "FFFFFF"
_DEFAULT_MUTED = "404040"
_DEFAULT_TITLE_ROW = "064A9C"
_DEFAULT_OK_GREEN = "006100"


def _hex_to_rgb(h: str) -> Tuple[int, int, int]:
    h = h.strip().lstrip("#")
    if len(h) != 6:
        return (0, 0, 0)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _rgb_to_hex_tuple(r: int, g: int, b: int) -> str:
    return f"{max(0, min(255, r)):02X}{max(0, min(255, g)):02X}{max(0, min(255, b)):02X}"


def _blend_hex(base_hex: str, tint_hex: str, amount: float) -> str:
    """Linear blend: result = (1-amount)*base + amount*tint."""
    a = max(0.0, min(1.0, amount))
    br, bg, bb = _hex_to_rgb(base_hex)
    tr, tg, tb = _hex_to_rgb(tint_hex)
    r = int(br * (1 - a) + tr * a)
    g = int(bg * (1 - a) + tg * a)
    b = int(bb * (1 - a) + tb * a)
    return _rgb_to_hex_tuple(r, g, b)


def _is_neutral_gray(rgb: Tuple[int, int, int]) -> bool:
    r, g, b = rgb
    spread = max(r, g, b) - min(r, g, b)
    return spread < 28


def _accent_colors_from_blocks(blocks: List["TextBlock"]) -> List[Tuple[int, int, int]]:
    """Non-gray text colors (often blues/reds in government PDFs)."""
    out: List[Tuple[int, int, int]] = []
    for b in blocks:
        if not (b.text or "").strip():
            continue
        rgb = b.color_rgb
        if _is_neutral_gray(rgb):
            continue
        # Skip near-white (header text on colored band)
        if sum(rgb) > 720:
            continue
        out.append(rgb)
    return out


@dataclass(frozen=True)
class SheetTheme:
    """PDF-derived typography + palette for one worksheet."""

    body_pt: float
    header_pt: float
    title_pt: float
    small_pt: float
    label_pt: float
    row_scale: float
    primary_fill_hex: str
    sub_fill_hex: str
    master_fill_hex: str
    title_fill_hex: str
    data_zebra_hex: str
    body_font_hex: str
    header_on_primary_hex: str
    muted_font_hex: str
    title_row_font_hex: str
    ok_green_hex: str
    landscape: bool
    paper_size: Optional[int]  # Excel paper size code (1=Letter, 9=A4, …)


def default_sheet_theme() -> SheetTheme:
    """Legacy constants (11pt body, TRACES blues)."""
    return SheetTheme(
        body_pt=11.0,
        header_pt=11.0,
        title_pt=14.0,
        small_pt=9.0,
        label_pt=10.0,
        row_scale=1.0,
        primary_fill_hex=_DEFAULT_PRIMARY,
        sub_fill_hex=_DEFAULT_SUB,
        master_fill_hex=_DEFAULT_MASTER,
        title_fill_hex=_DEFAULT_TITLE,
        data_zebra_hex=_DEFAULT_ZEBRA,
        body_font_hex=_DEFAULT_BODY_FONT,
        header_on_primary_hex=_DEFAULT_HEADER_ON_PRIMARY,
        muted_font_hex=_DEFAULT_MUTED,
        title_row_font_hex=_DEFAULT_TITLE_ROW,
        ok_green_hex=_DEFAULT_OK_GREEN,
        landscape=False,
        paper_size=None,
    )


def _median_size(blocks: List["TextBlock"]) -> float:
    sizes = [float(b.size) for b in blocks if (b.text or "").strip()]
    if not sizes:
        return 11.0
    return float(median(sizes))


def _page_dimensions(parsed: "ParsedPDF", page_idx: int) -> Tuple[float, float]:
    pw = float(parsed.page_widths[page_idx]) if page_idx < len(parsed.page_widths) else 612.0
    ph = float(parsed.page_heights[page_idx]) if page_idx < len(parsed.page_heights) else 792.0
    return pw, ph


def guess_landscape(parsed: "ParsedPDF", page_idx: int) -> bool:
    pw, ph = _page_dimensions(parsed, page_idx)
    return pw > ph


def guess_paper_size(parsed: "ParsedPDF", page_idx: int) -> int:
    """
    Map PDF page size in points to Excel paper size.
    9 = A4, 1 = Letter (common openpyxl / Excel mappings).
    """
    pw, ph = _page_dimensions(parsed, page_idx)
    w, h = min(pw, ph), max(pw, ph)
    # A4 portrait ~ 595 x 842
    if abs(w - 595) < 50 and abs(h - 842) < 50:
        return 9
    # US Letter ~ 612 x 792
    if abs(w - 612) < 50 and abs(h - 792) < 50:
        return 1
    # Legal / other → A4 as safe default for metric PDFs
    if w < 600:
        return 9
    return 1


def build_theme_for_page(
    parsed: "ParsedPDF",
    page_idx: int,
    *,
    is_traces: bool = False,
) -> SheetTheme:
    """
    Aggregate font sizes from TextBlock on this page; lightly tint header fills
    toward non-gray PDF text colors when present.
    """
    base = default_sheet_theme()
    blocks = [b for b in parsed.text_blocks if getattr(b, "page", -1) == page_idx]
    if not blocks:
        return SheetTheme(
            body_pt=base.body_pt,
            header_pt=base.header_pt,
            title_pt=base.title_pt,
            small_pt=base.small_pt,
            label_pt=base.label_pt,
            row_scale=base.row_scale,
            primary_fill_hex=base.primary_fill_hex,
            sub_fill_hex=base.sub_fill_hex,
            master_fill_hex=base.master_fill_hex,
            title_fill_hex=base.title_fill_hex,
            data_zebra_hex=base.data_zebra_hex,
            body_font_hex=base.body_font_hex,
            header_on_primary_hex=base.header_on_primary_hex,
            muted_font_hex=base.muted_font_hex,
            title_row_font_hex=base.title_row_font_hex,
            ok_green_hex=base.ok_green_hex,
            landscape=guess_landscape(parsed, page_idx),
            paper_size=guess_paper_size(parsed, page_idx),
        )

    body_raw = _median_size(blocks)
    body_pt = max(8.0, min(14.0, round(body_raw)))

    bold_sizes = [float(b.size) for b in blocks if b.bold and (b.text or "").strip()]
    if bold_sizes:
        header_pt = max(body_pt, min(16.0, round(max(bold_sizes))))
    else:
        header_pt = min(16.0, body_pt + 1.0)

    title_pt = min(18.0, max(header_pt + 1.0, body_pt + 3.0))
    small_pt = max(8.0, body_pt - 2.0)
    label_pt = max(8.0, min(12.0, body_pt - 1.0))

    row_scale = body_pt / 11.0

    accents = _accent_colors_from_blocks(blocks)
    tint_amount = 0.12 if is_traces else 0.18
    if accents:
        ar = sum(a[0] for a in accents) // len(accents)
        ag = sum(a[1] for a in accents) // len(accents)
        ab = sum(a[2] for a in accents) // len(accents)
        tint_hex = _rgb_to_hex_tuple(ar, ag, ab)
        primary = _blend_hex(_DEFAULT_PRIMARY, tint_hex, tint_amount)
        sub = _blend_hex(_DEFAULT_SUB, tint_hex, tint_amount * 0.5)
        master = _blend_hex(_DEFAULT_MASTER, tint_hex, tint_amount * 0.35)
        title_f = _blend_hex(_DEFAULT_TITLE, tint_hex, tint_amount * 0.6)
    else:
        primary, sub, master, title_f = (
            _DEFAULT_PRIMARY,
            _DEFAULT_SUB,
            _DEFAULT_MASTER,
            _DEFAULT_TITLE,
        )

    # Body font color: median of dark text (avoid white header fragments)
    dark_hexes = [
        rgb_to_hex(b.color_rgb)
        for b in blocks
        if (b.text or "").strip() and sum(b.color_rgb) < 500
    ]
    if dark_hexes:
        # use first dark color as representative (stable enough)
        body_font_hex = dark_hexes[len(dark_hexes) // 2]
    else:
        body_font_hex = _DEFAULT_BODY_FONT

    return SheetTheme(
        body_pt=body_pt,
        header_pt=header_pt,
        title_pt=title_pt,
        small_pt=small_pt,
        label_pt=label_pt,
        row_scale=row_scale,
        primary_fill_hex=primary,
        sub_fill_hex=sub,
        master_fill_hex=master,
        title_fill_hex=title_f,
        data_zebra_hex=_DEFAULT_ZEBRA,
        body_font_hex=body_font_hex,
        header_on_primary_hex=_DEFAULT_HEADER_ON_PRIMARY,
        muted_font_hex=_DEFAULT_MUTED,
        title_row_font_hex=_DEFAULT_TITLE_ROW,
        ok_green_hex=_DEFAULT_OK_GREEN,
        landscape=guess_landscape(parsed, page_idx),
        paper_size=guess_paper_size(parsed, page_idx),
    )


def apply_worksheet_page_setup(ws: Any, theme: SheetTheme) -> None:
    """
    Match PDF Page sheets: fit print to one page wide/tall, margins, orientation, paper.
    """
    try:
        from openpyxl.worksheet.page import PageMargins
    except ImportError:
        return
    try:
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    try:
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
    except Exception:
        pass
    try:
        ws.page_margins = PageMargins(
            left=0.2, right=0.2, top=0.25, bottom=0.25, header=0.0, footer=0.0
        )
    except Exception:
        pass
    try:
        ws.page_setup.orientation = "landscape" if theme.landscape else "portrait"
    except Exception:
        pass
    if theme.paper_size is not None:
        try:
            ws.page_setup.paperSize = theme.paper_size
        except Exception:
            pass
