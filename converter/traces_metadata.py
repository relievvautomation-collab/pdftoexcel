"""Detect TRACES / Form 26AS-style PDFs and extract assessee header fields from page-0 text."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from .pdf_parser import ParsedPDF, TextBlock

_PAN_RE = re.compile(r"\b([A-Z]{5}[0-9]{4}[A-Z])\b")
_FY_RE = re.compile(
    r"(?:Financial\s+Year|F\.?\s*Y\.?)\s*[:]?\s*(\d{4}-\d{2})",
    re.I,
)
_AY_RE = re.compile(
    r"(?:Assessment\s+Year|A\.?\s*Y\.?)\s*[:]?\s*(\d{4}-\d{2})",
    re.I,
)
_FY_AY_LINE_RE = re.compile(
    r"Financial\s+Year\s*[:]?\s*(\d{4}-\d{2}).{0,120}?Assessment\s+Year\s*[:]?\s*(\d{4}-\d{2})",
    re.I | re.DOTALL,
)
_UPDATED_RE = re.compile(
    r"Data\s+updated\s+till\s*[:]?\s*([^\n]+?)(?:\s{2,}|\n|$)",
    re.I,
)
_PAN_STATUS_RE = re.compile(
    r"Current\s+Status\s+of\s+PAN\s*[:]?\s*([^\n]+?)(?:\n|Financial|\s{3,}|$)",
    re.I,
)
_NAME_RE = re.compile(
    r"Name\s+of\s+Assessee\s*[:]?\s*([^\n]+?)(?:\n|Address|\s{3,}|$)",
    re.I,
)
_ADDR_RE = re.compile(
    r"Address\s+of\s+Assessee\s*[:]?\s*(.+?)(?:\n\s*\n|Part-|$)",
    re.I | re.DOTALL,
)


def _page0_text_blocks(blocks: List["TextBlock"]) -> List["TextBlock"]:
    return sorted(
        [b for b in blocks if getattr(b, "page", -1) == 0],
        key=lambda b: (round(b.y0, 2), round(b.x0, 2)),
    )


def page0_plain_text(parsed: "ParsedPDF") -> str:
    """Single string for page 0 (reading order)."""
    parts: List[str] = []
    prev_y: Optional[float] = None
    for b in _page0_text_blocks(parsed.text_blocks):
        if prev_y is not None and abs(b.y0 - prev_y) > 5:
            parts.append("\n")
        parts.append(b.text.strip())
        prev_y = b.y0
    return "\n".join(parts) if parts else " ".join(b.text for b in _page0_text_blocks(parsed.text_blocks))


def detect_traces_26as(parsed: "ParsedPDF") -> bool:
    """
    Heuristic: TRACES Annual Tax Statement / Part-I TDS PDFs.
    """
    t = page0_plain_text(parsed).lower()
    if len(t) < 80:
        return False
    score = 0
    if "annual tax statement" in t:
        score += 3
    if "traces" in t and ("tds" in t or "reconciliation" in t):
        score += 2
    if "part-i" in t or "part i" in t:
        score += 2
    if "name of deductor" in t and "tan of deductor" in t:
        score += 3
    if "form 26as" in t or "26as" in t:
        score += 2
    if "income tax" in t and "department" in t:
        score += 1
    if _PAN_RE.search(t) and ("assessee" in t or "permanent account" in t):
        score += 1
    if "total tds" in t or "total tax deducted" in t:
        score += 1
    return score >= 3


def detect_traces_from_extracted_tables(parsed: "ParsedPDF") -> bool:
    """
    Infer TRACES Part-I TDS when page-0 wording is missing or OCR is noisy, but
    Camelot captured the standard deductor / detail grid.
    """
    for tab in getattr(parsed, "camelot_tables", None) or []:
        data = tab.get("data") or []
        if not data:
            continue
        sample = " ".join(
            " ".join(str(c) for c in row if c is not None and str(c).strip())
            for row in data[:30]
        )
        lo = sample.lower()
        hits = 0
        if "name of deductor" in lo or ("deductor" in lo and "tan" in lo):
            hits += 2
        if _PAN_RE.search(sample) and "assessee" in lo:
            hits += 1
        if re.search(r"\b[a-z]{4}\d{5}[a-z]\b", sample, re.I):
            hits += 2
        if "section" in lo and (
            "transaction" in lo or "remarks" in lo or "booking" in lo
        ):
            hits += 1
        if "total tds" in lo or "total tax deducted" in lo or "amount paid" in lo:
            hits += 1
        if hits >= 3:
            return True
    return False


def is_traces_pdf(parsed: "ParsedPDF") -> bool:
    """True if this workbook should use TRACES / 26AS Part-I layout and normalization."""
    return bool(detect_traces_26as(parsed) or detect_traces_from_extracted_tables(parsed))


def extract_assessee_header(parsed: "ParsedPDF") -> Dict[str, Any]:
    """
    Best-effort fields for the assessee summary block (from page-0 text, not Camelot).
    """
    raw = page0_plain_text(parsed)
    # Flatten for regex (also try single-line variant)
    flat = " ".join(raw.split())

    out: Dict[str, Any] = {
        "pan": "",
        "pan_status": "",
        "financial_year": "",
        "assessment_year": "",
        "name": "",
        "address": "",
        "data_updated_till": "",
    }

    m = _PAN_RE.search(raw) or _PAN_RE.search(flat)
    if m:
        out["pan"] = m.group(1).upper()

    m = _UPDATED_RE.search(raw)
    if m:
        out["data_updated_till"] = m.group(1).strip()

    m = _FY_RE.search(raw) or _FY_RE.search(flat)
    if m:
        out["financial_year"] = m.group(1).strip()

    m = _AY_RE.search(raw) or _AY_RE.search(flat)
    if m:
        out["assessment_year"] = m.group(1).strip()

    m = _FY_AY_LINE_RE.search(raw) or _FY_AY_LINE_RE.search(flat)
    if m:
        if not out["financial_year"]:
            out["financial_year"] = m.group(1).strip()
        if not out["assessment_year"]:
            out["assessment_year"] = m.group(2).strip()

    m = _PAN_STATUS_RE.search(raw)
    if m:
        out["pan_status"] = m.group(1).strip()[:200]

    m = _NAME_RE.search(raw)
    if m:
        out["name"] = m.group(1).strip()[:500]

    m = _ADDR_RE.search(raw)
    if m:
        addr = re.sub(r"\s+", " ", m.group(1).strip())
        out["address"] = addr[:1000]

    # Fallback PAN-only line: "PAN : XXXXX"
    if not out["pan"]:
        m2 = re.search(
            r"(?:Permanent\s+Account\s+Number\s*\(?PAN\)?|PAN)\s*[:]?\s*([A-Z]{5}[0-9]{4}[A-Z])",
            raw,
            re.I,
        )
        if m2:
            out["pan"] = m2.group(1).upper()

    # Fallback FY/AY: two year tokens like 2024-25 / 2025-26 on one line of page text
    if not out["financial_year"] or not out["assessment_year"]:
        years = re.findall(r"\b(20\d{2}-\d{2})\b", flat)
        if len(years) >= 1 and not out["financial_year"]:
            out["financial_year"] = years[0]
        if len(years) >= 2 and not out["assessment_year"]:
            out["assessment_year"] = years[1]

    return out


_DEFAULT_PART_I = "PART-I - Details of Tax Deducted at Source"
_DEFAULT_INR = "(All amount values are in INR)"

# PAN disclaimer: from "Above data" until PART-I / Part-I block or end of paragraph
_PAN_DISCLAIMER_RE = re.compile(
    r"(Above\s+data[\s\S]{10,2000}?)(?=\s*(?:PART\s*-?\s*I|Part\s*-?\s*I|Details\s+of\s+Tax\s+Deducted|\Z))",
    re.I,
)
_PART_I_TITLE_RE = re.compile(
    r"(PART\s*-?\s*I\s*[-—–]\s*Details\s+of\s+Tax\s+Deducted\s+at\s+Source)",
    re.I,
)
_INR_NOTE_RE = re.compile(
    r"(\(?All\s+amount\s+values?\s+are\s+in\s+INR\)?)",
    re.I,
)


def extract_traces_preamble(parsed: "ParsedPDF") -> Dict[str, str]:
    """
    Disclaimer, Part-I section title, and INR note from page-0 text (TRACES Annual Tax Statement).
    Safe defaults when OCR/layout splits text unpredictably.
    """
    raw = page0_plain_text(parsed)
    flat = " ".join(raw.split())

    pan_disclaimer = ""
    m = _PAN_DISCLAIMER_RE.search(raw)
    if m:
        pan_disclaimer = re.sub(r"\s+", " ", m.group(1).strip())[:1200]
    else:
        m2 = re.search(
            r"(Above\s+data[\s\S]{10,1500}?)(?:tin-nsdl|utiitsl|PART\s*-?\s*I)",
            raw,
            re.I,
        )
        if m2:
            pan_disclaimer = re.sub(r"\s+", " ", m2.group(1).strip())[:1200]

    part_i_title = _DEFAULT_PART_I
    m = _PART_I_TITLE_RE.search(raw) or _PART_I_TITLE_RE.search(flat)
    if m:
        part_i_title = re.sub(r"\s+", " ", m.group(1).strip())[:200]

    inr_note = _DEFAULT_INR
    m = _INR_NOTE_RE.search(raw) or _INR_NOTE_RE.search(flat)
    if m:
        inr_note = m.group(1).strip()[:120]

    return {
        "pan_disclaimer": pan_disclaimer,
        "part_i_title": part_i_title,
        "inr_note": inr_note,
    }
