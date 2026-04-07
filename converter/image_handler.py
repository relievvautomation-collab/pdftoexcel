"""Extract embedded images from PDF pages (PyMuPDF)."""

from __future__ import annotations

import io
from dataclasses import dataclass
from typing import List, Optional

import fitz


@dataclass
class ExtractedImage:
    page_index: int
    xref: int
    bbox: tuple  # (x0, y0, x1, y1) in PDF points
    width_px: int
    height_px: int
    ext: str
    data: bytes


def extract_images_from_pdf(path: str) -> List[ExtractedImage]:
    doc = fitz.open(path)
    out: List[ExtractedImage] = []
    seen: set = set()
    try:
        for pno in range(len(doc)):
            page = doc[pno]
            for img in page.get_images(full=True):
                xref = img[0]
                key = (pno, xref)
                if key in seen:
                    continue
                seen.add(key)
                try:
                    base = doc.extract_image(xref)
                except Exception:
                    continue
                rects = page.get_image_rects(xref)
                bbox = (0.0, 0.0, float(page.rect.width), float(page.rect.height))
                if rects:
                    r = rects[0]
                    bbox = (r.x0, r.y0, r.x1, r.y1)
                out.append(
                    ExtractedImage(
                        page_index=pno,
                        xref=xref,
                        bbox=bbox,
                        width_px=base.get("width", 0),
                        height_px=base.get("height", 0),
                        ext=base.get("ext", "png"),
                        data=base.get("image", b""),
                    )
                )
    finally:
        doc.close()
    return out


def image_to_png_bytes(data: bytes, ext: str) -> bytes:
    """Ensure PNG bytes for openpyxl."""
    if ext == "png" or not data:
        return data
    try:
        from PIL import Image

        im = Image.open(io.BytesIO(data))
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGBA")
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return data
