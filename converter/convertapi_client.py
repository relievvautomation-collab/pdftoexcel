"""ConvertAPI PDF → XLSX using the official Python client (https://github.com/ConvertAPI/convertapi-python)."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict

import convertapi

# POST https://v2.convertapi.com/convert/pdf/to/xlsx
# OpenAPI: SingleSheet, IncludeFormatting — see https://v2.convertapi.com/info/openapi/pdf/to/xlsx


def convert_pdf_to_xlsx_convertapi(
    pdf_path: Path,
    out_xlsx_path: Path,
    secret: str,
    *,
    timeout_sec: int = 600,
    single_sheet: bool = False,
    include_formatting: bool = False,
) -> Dict[str, Any]:
    """
    Convert PDF to XLSX via ConvertAPI.

    - ``single_sheet``: maps to ``SingleSheet`` — combine tables into one worksheet.
    - ``include_formatting``: maps to ``IncludeFormatting`` — include images / non-table text when supported.
    """
    secret = (secret or "").strip()
    if not secret:
        raise ValueError("ConvertAPI secret is empty; set CONVERTAPI_SECRET")

    pdf_path = Path(pdf_path).resolve()
    out_xlsx_path = Path(out_xlsx_path)
    if not pdf_path.is_file():
        raise FileNotFoundError(str(pdf_path))

    out_xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    convertapi.api_credentials = secret

    base = (os.environ.get("CONVERTAPI_BASE_URI") or "").strip()
    if base:
        convertapi.base_uri = base.rstrip("/") + "/"

    params: Dict[str, Any] = {"File": str(pdf_path)}
    if single_sheet:
        params["SingleSheet"] = True
    if include_formatting:
        params["IncludeFormatting"] = True

    result = convertapi.convert(
        "xlsx",
        params,
        from_format="pdf",
        timeout=timeout_sec,
    )

    result.file.save(str(out_xlsx_path))

    out: Dict[str, Any] = {
        "conversion_cost": result.conversion_cost,
        "file_name": result.file.filename,
        "file_size": result.file.size,
        "single_sheet": single_sheet,
        "include_formatting": include_formatting,
    }
    return out
