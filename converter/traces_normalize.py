"""Normalize TRACES Part-I detail rows when Camelot merges Section + Transaction Date into one cell."""

from __future__ import annotations

import re
from typing import List, Optional

from .statement_excel import _is_sub_header, _join_row

# Section code + date: spaced or hyphenated (e.g. 12 Sep 2024, 12-Sep-2024)
_SPLIT_SECTION_DATE = re.compile(
    r"^\s*(?:Section\s+)?(?P<section>19\d{1,2}[A-Z]?)\s+"
    r"(?P<date>\d{1,2}[\s\-]+[A-Za-z]{3,9}[\s\-]+\d{2,4})\s*$",
    re.I,
)
# "194 | 12-Sep-2024" or "194|12 Sep 2024" (common in exports)
_SPLIT_PIPE_SECTION_DATE = re.compile(
    r"^\s*(?:Section\s+)?(?P<section>19\d{1,2}[A-Z]?)\s*\|\s*(?P<date>.+?)\s*$",
    re.I,
)


TRACES_DETAIL_TARGET_COLS = 9  # Sr, Section, Txn date, Status, Booking date, Remarks, 3 amounts


def pad_traces_detail_to_standard_width(data: List[List[str]]) -> List[List[str]]:
    """
    TRACES Part-I detail grids are typically 9 columns; Camelot often yields 6–8.
    Pad rows on the right with empty cells so headers and amounts align like the PDF.
    """
    if not data:
        return data
    sub_i = _find_sub_header_row(data)
    if sub_i is None:
        return data
    max_c = max(len(r) for r in data)
    joined = _join_row(data[sub_i])
    if "section" not in joined:
        return data
    if not any(x in joined for x in ("amount", "tds", "tax", "credited", "remarks")):
        return data
    target = max(max_c, TRACES_DETAIL_TARGET_COLS)
    return [r + [""] * (target - len(r)) for r in data]


def _find_sub_header_row(data: List[List[str]]) -> Optional[int]:
    for i, row in enumerate(data):
        if _is_sub_header(_join_row(row)):
            return i
    return None


def _header_cell_suggests_section_transaction_merge(row: List[str]) -> Optional[int]:
    """Return column index where header looks like merged Section / Transaction (not separate Transaction Date)."""
    for j, cell in enumerate(row):
        t = (cell or "").strip().lower()
        if not t:
            continue
        if "section" in t and "transaction" in t and "/" in t:
            return j
        if t == "section / transaction" or "section/transaction" in t.replace(" ", ""):
            return j
        if "section" in t and "transaction" in t and "date" not in t:
            return j
        # OCR: "Section 1 Transaction" as one header
        if re.search(r"section\s*\d*\s*transaction", t, re.I) and "date" not in t:
            return j
    return None


def normalize_traces_table_columns(data: List[List[str]]) -> List[List[str]]:
    """
    If detail header has a merged Section/Transaction column and body cells match
    section+date pattern, split into two columns at that index (pads all rows).
    """
    if not data:
        return data
    sub_i = _find_sub_header_row(data)
    if sub_i is None:
        return pad_traces_detail_to_standard_width([list(r) for r in data])
    col = _header_cell_suggests_section_transaction_merge(data[sub_i])
    if col is None:
        return pad_traces_detail_to_standard_width([list(r) for r in data])

    new_rows: List[List[str]] = []
    for ri, row in enumerate(data):
        r = list(row)
        if ri == sub_i:
            if col < len(r):
                r[col : col + 1] = ["Section", "Transaction Date"]
        elif col < len(r):
            cell = (r[col] or "").strip()
            m = _SPLIT_SECTION_DATE.match(cell) or _SPLIT_PIPE_SECTION_DATE.match(cell)
            if m:
                sec = m.group("section")
                if not sec.lower().startswith("section"):
                    sec = f"Section {sec}"
                dt = m.group("date").strip()
                r[col : col + 1] = [sec, dt]
        new_rows.append(r)

    max_c = max(len(x) for x in new_rows) if new_rows else 0
    padded = [x + [""] * (max_c - len(x)) for x in new_rows]
    return pad_traces_detail_to_standard_width(padded)


def find_status_of_booking_column(data: List[List[str]]) -> Optional[int]:
    """0-based column index for 'Status of Booking' in sub-header, if present."""
    sub_i = _find_sub_header_row(data)
    if sub_i is None:
        return None
    row = data[sub_i]
    for j, cell in enumerate(row):
        t = (cell or "").strip().lower()
        if "status" in t and "booking" in t:
            return j
    return None
