"""Extract text, images, styles, drawings, and optional tables from PDF."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import fitz
import numpy as np

from .color_utils import fitz_color_to_rgb, int_to_rgb


@dataclass
class TextBlock:
    text: str
    page: int
    x0: float
    y0: float
    x1: float
    y1: float
    font: str
    size: float
    color_rgb: Tuple[int, int, int]
    bold: bool
    italic: bool
    underline: bool
    align: str  # left, center, right


@dataclass
class LineDrawing:
    page: int
    x0: float
    y0: float
    x1: float
    y1: float
    width: float
    color_rgb: Tuple[int, int, int]
    dashed: bool


@dataclass
class ParsedPDF:
    page_widths: List[float]
    page_heights: List[float]
    text_blocks: List[TextBlock]
    images: List[Any]  # ExtractedImage from image_handler
    drawings: List[LineDrawing]
    camelot_tables: List[Dict[str, Any]]
    plumber_tables: List[Dict[str, Any]]
    used_ocr: bool
    meta: Dict[str, Any] = field(default_factory=dict)


def _span_flags_to_style(flags: int) -> Tuple[bool, bool, bool]:
    # PyMuPDF span flags: see TextPage — bold ~ bit 3, italic ~ bit 1
    bold = bool(flags & (1 << 3))
    italic = bool(flags & (1 << 1))
    underline = bool(flags & (1 << 2))  # approximate; PDFs vary
    return bold, italic, underline


def _guess_align(page_width: float, x0: float, x1: float, text: str) -> str:
    mid = (x0 + x1) / 2
    if mid < page_width * 0.35:
        return "left"
    if mid > page_width * 0.65:
        return "right"
    return "center"


def extract_text_blocks(doc: fitz.Document, ocr_if_empty: bool = True) -> Tuple[List[TextBlock], bool]:
    blocks: List[TextBlock] = []
    used_ocr = False
    for pno in range(len(doc)):
        page = doc[pno]
        page_w = float(page.rect.width)
        d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_PRESERVE_LIGATURES)
        page_has_text = False
        for b in d.get("blocks", []):
            if b.get("type") != 0:
                continue
            for line in b.get("lines", []):
                for span in line.get("spans", []):
                    txt = (span.get("text") or "").strip()
                    if not txt:
                        continue
                    page_has_text = True
                    bbox = span.get("bbox", [0, 0, 0, 0])
                    x0, y0, x1, y1 = bbox
                    color = span.get("color", 0)
                    if isinstance(color, float):
                        color = int(color)
                    rgb = int_to_rgb(int(color)) if isinstance(color, int) else fitz_color_to_rgb(color)
                    flags = int(span.get("flags", 0))
                    bold, italic, underline = _span_flags_to_style(flags)
                    font = str(span.get("font", "Calibri"))
                    size = float(span.get("size", 11))
                    align = _guess_align(page_w, x0, x1, txt)
                    blocks.append(
                        TextBlock(
                            text=txt,
                            page=pno,
                            x0=x0,
                            y0=y0,
                            x1=x1,
                            y1=y1,
                            font=font,
                            size=size,
                            color_rgb=rgb,
                            bold=bold,
                            italic=italic,
                            underline=underline,
                            align=align,
                        )
                    )
        if ocr_if_empty and not page_has_text:
            try:
                import pytesseract
                from PIL import Image

                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
                n = len(data.get("text", []))
                for i in range(n):
                    w = str(data["text"][i] or "").strip()
                    if not w:
                        continue
                    x, y, ww, hh = (
                        data["left"][i],
                        data["top"][i],
                        data["width"][i],
                        data["height"][i],
                    )
                    # scale back to page coords (approximate)
                    sx = float(page.rect.width) / pix.width
                    sy = float(page.rect.height) / pix.height
                    x0 = x * sx
                    y0 = y * sy
                    x1 = (x + ww) * sx
                    y1 = (y + hh) * sy
                    blocks.append(
                        TextBlock(
                            text=w,
                            page=pno,
                            x0=x0,
                            y0=y0,
                            x1=x1,
                            y1=y1,
                            font="Arial",
                            size=float(max(8, hh * sy * 0.75)),
                            color_rgb=(0, 0, 0),
                            bold=False,
                            italic=False,
                            underline=False,
                            align="left",
                        )
                    )
                used_ocr = True
            except Exception:
                pass
    return blocks, used_ocr


def extract_drawings(doc: fitz.Document) -> List[LineDrawing]:
    lines: List[LineDrawing] = []
    for pno in range(len(doc)):
        page = doc[pno]
        try:
            for d in page.get_drawings():
                rect = d.get("rect")
                if rect is None:
                    continue
                color = d.get("color")
                rgb = fitz_color_to_rgb(color) if color is not None else (200, 200, 200)
                width = float(d.get("width", 0.5) or 0.5)
                dashed = d.get("dashes") not in (None, "")
                # Normalize line segments
                items = d.get("items", [])
                for it in items:
                    if it[0] == "l" and len(it) >= 3:
                        p1, p2 = it[1], it[2]
                        lines.append(
                            LineDrawing(
                                page=pno,
                                x0=float(p1.x),
                                y0=float(p1.y),
                                x1=float(p2.x),
                                y1=float(p2.y),
                                width=width,
                                color_rgb=rgb,
                                dashed=dashed,
                            )
                        )
        except Exception:
            continue
    return lines


def _normalize_table_rows(raw: List[List[Any]]) -> List[List[str]]:
    out: List[List[str]] = []
    for row in raw or []:
        cells = [("" if c is None else str(c)).strip() for c in row]
        out.append(cells)
    return out


def _table_nonempty(data: List[List[str]]) -> bool:
    return any(any(c for c in row) for row in data)


def run_pdfplumber_tables(path: str) -> List[Dict[str, Any]]:
    """Detect bordered / text-aligned tables for editable Excel export."""
    tables_out: List[Dict[str, Any]] = []
    try:
        import pdfplumber

        strategies: List[Dict[str, Any]] = [
            {},
            {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
            {
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "intersection_tolerance": 5,
            },
        ]
        with pdfplumber.open(path) as pdf:
            for pi, page in enumerate(pdf.pages):
                seen: set = set()
                for strat in strategies:
                    try:
                        raw_tbls = page.extract_tables(table_settings=strat) or []
                    except Exception:
                        raw_tbls = []
                    for tbl in raw_tbls:
                        data = _normalize_table_rows(tbl)
                        if not _table_nonempty(data):
                            continue
                        sig = (pi, len(data), len(data[0]) if data else 0, tuple(data[0][:6]))
                        if sig in seen:
                            continue
                        seen.add(sig)
                        tables_out.append(
                            {
                                "page": pi,
                                "source": "pdfplumber",
                                "strategy": str(strat),
                                "data": data,
                            }
                        )
    except Exception:
        pass
    return tables_out


def run_camelot_tables(path: str) -> List[Dict[str, Any]]:
    tables: List[Dict[str, Any]] = []
    try:
        import camelot

        for flavor in ("lattice", "stream"):
            try:
                tlist = camelot.read_pdf(path, pages="all", flavor=flavor)
                for t in tlist:
                    df = t.df
                    tables.append(
                        {
                            "page": int(t.page) - 1,
                            "flavor": flavor,
                            "accuracy": float(t.accuracy),
                            "data": df.replace({np.nan: ""}).values.tolist(),
                            "shape": list(df.shape),
                            "_bbox": getattr(t, "_bbox", None),
                        }
                    )
                if tables:
                    break
            except Exception:
                continue
    except Exception:
        pass
    return tables


def parse_pdf(path: str) -> ParsedPDF:
    from .image_handler import extract_images_from_pdf

    doc = fitz.open(path)
    try:
        widths = [float(doc[p].rect.width) for p in range(len(doc))]
        heights = [float(doc[p].rect.height) for p in range(len(doc))]
        text_blocks, used_ocr = extract_text_blocks(doc, ocr_if_empty=True)
        drawings = extract_drawings(doc)
        images = extract_images_from_pdf(path)
        camelot_tables = run_camelot_tables(path)
        plumber_tables = run_pdfplumber_tables(path)
        meta = {"page_count": len(doc)}
    finally:
        doc.close()
    return ParsedPDF(
        page_widths=widths,
        page_heights=heights,
        text_blocks=text_blocks,
        images=images,
        drawings=drawings,
        camelot_tables=camelot_tables,
        plumber_tables=plumber_tables,
        used_ocr=used_ocr,
        meta=meta,
    )
