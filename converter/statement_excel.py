"""Heuristic row/column styling for tax-statement-style PDF tables (e.g. TRACES 26AS)."""

from __future__ import annotations

import re
from enum import Enum
from typing import List, Optional


class RowKind(str, Enum):
    TITLE = "title"
    PRIMARY_HEADER = "primary_header"
    SUB_HEADER = "sub_header"
    MASTER_SUMMARY = "master_summary"
    DATA = "data"
    FALLBACK_HEADER = "fallback_header"  # first row when no keyword match


def _norm_cell(s: str) -> str:
    return (s or "").strip().lower()


def _join_row(row: List[str]) -> str:
    return " ".join(_norm_cell(c) for c in row if c is not None)


def _count_numeric_cells(row: List[str]) -> int:
    n = 0
    for c in row:
        if try_parse_number(str(c) if c is not None else "") is not None:
            n += 1
    return n


# Indian TAN: 4 letters + 5 digits + 1 letter
_TAN_RE = re.compile(r"\b[a-z]{4}\d{5}[a-z]\b", re.I)


def try_parse_number(s: str) -> Optional[float]:
    t = str(s).strip().replace(",", "").replace("−", "-").replace("—", "-")
    if not t or t in ("-", "–"):
        return None
    try:
        return float(t)
    except ValueError:
        return None


def _is_primary_header(joined: str) -> bool:
    if "part-" in joined and "detail" in joined:
        return True
    hits = 0
    for kw in (
        "name of deductor",
        "tan of deductor",
        "total amount paid",
        "total tax deducted",
        "total tds deposited",
        "amount paid / credited",
    ):
        if kw in joined:
            hits += 1
    if hits >= 2:
        return True
    if "sr" in joined and "deductor" in joined and "tan" in joined:
        return True
    return False


def _is_sub_header(joined: str) -> bool:
    """Detail grid under each deductor: Section + transaction/remarks/amounts (PDFs vary)."""
    if "section" not in joined:
        return False
    if "transaction" in joined and (
        "remarks" in joined
        or "amount paid" in joined
        or "tax deducted" in joined
        or "tds deposited" in joined
    ):
        return True
    return any(
        x in joined
        for x in (
            "transaction date",
            "status of booking",
            "date of booking",
            "remarks",
            "tax deducted",
            "tds deposited",
        )
    )


def _is_title_row(joined: str) -> bool:
    return ("page" in joined and ("camelot" in joined or "pdfplumber" in joined or "lattice" in joined)) or joined.startswith(
        "page "
    )


def _is_master_summary(row: List[str], joined: str) -> bool:
    if _TAN_RE.search(joined):
        return True
    nums = _count_numeric_cells(row)
    if nums >= 2:
        # Long text chunk typical of company name
        for c in row:
            t = str(c or "").strip()
            if len(t) > 18 and not try_parse_number(t):
                return True
    return False


def classify_statement_rows(data: List[List[str]]) -> List[RowKind]:
    """Assign a RowKind per data row for styling."""
    if not data:
        return []
    kinds: List[RowKind] = []
    prev: Optional[RowKind] = None

    for i, row in enumerate(data):
        joined = _join_row(row)

        if _is_title_row(joined):
            kinds.append(RowKind.TITLE)
            prev = RowKind.TITLE
            continue

        if _is_primary_header(joined):
            kinds.append(RowKind.PRIMARY_HEADER)
            prev = RowKind.PRIMARY_HEADER
            continue

        if _is_sub_header(joined):
            kinds.append(RowKind.SUB_HEADER)
            prev = RowKind.SUB_HEADER
            continue

        if prev in (RowKind.PRIMARY_HEADER, RowKind.TITLE, RowKind.FALLBACK_HEADER) and _is_master_summary(
            row, joined
        ):
            kinds.append(RowKind.MASTER_SUMMARY)
            prev = RowKind.MASTER_SUMMARY
            continue

        if prev == RowKind.SUB_HEADER:
            kinds.append(RowKind.DATA)
            prev = RowKind.DATA
            continue

        if i == 0:
            kinds.append(RowKind.FALLBACK_HEADER)
            prev = RowKind.FALLBACK_HEADER
            continue

        kinds.append(RowKind.DATA)
        prev = RowKind.DATA

    while len(kinds) < len(data):
        kinds.append(RowKind.DATA)
    return kinds[: len(data)]


def header_suggests_amount_cell(header_text: str) -> bool:
    t = _norm_cell(header_text)
    return any(
        x in t
        for x in (
            "amount",
            "tax",
            "tds",
            "paid",
            "credited",
            "deducted",
            "deposited",
            "total",
        )
    )


def normalize_header_fingerprint(row: List[str]) -> str:
    """Stable key for comparing repeated header rows from Camelot."""
    return "|".join(_norm_cell(str(c)) for c in row)


def dedupe_tds_extraction_rows(data: List[List[str]]) -> List[List[str]]:
    """
    Camelot often emits the same deductor header row before every block, which looks
    like overlapping / repeated headers in Excel. Drop duplicate header bands.
    """
    if not data:
        return data
    out: List[List[str]] = []
    primary_fp: Optional[str] = None
    sub_fp: Optional[str] = None

    for row in data:
        joined = _join_row(row)

        if _is_title_row(joined):
            out.append(row)
            continue

        if _is_primary_header(joined):
            fp = normalize_header_fingerprint(row)
            if primary_fp is None:
                primary_fp = fp
                out.append(row)
            elif fp == primary_fp:
                continue
            else:
                out.append(row)
            continue

        if _is_sub_header(joined):
            fp = normalize_header_fingerprint(row)
            if sub_fp is None:
                sub_fp = fp
                out.append(row)
            elif (
                fp == sub_fp
                and out
                and _is_sub_header(_join_row(out[-1]))
                and normalize_header_fingerprint(out[-1]) == fp
            ):
                continue
            else:
                out.append(row)
            continue

        out.append(row)

    return out


def build_amount_column_mask(header_rows: List[List[str]], ncols: int) -> List[bool]:
    """True = column should be right-aligned / number formatted when parseable."""
    mask = [False] * max(0, ncols)
    for hr in header_rows[:3]:
        for j, cell in enumerate(hr):
            if j >= ncols:
                break
            if header_suggests_amount_cell(str(cell)):
                mask[j] = True
    # Default: right-align last few columns if table is wide (typical amounts on right)
    if ncols >= 4 and not any(mask[-3:]):
        for j in range(max(0, ncols - 3), ncols):
            mask[j] = True
    return mask
