"""Flask application: PDF upload, conversion, preview, download."""

from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

try:
    from dotenv import load_dotenv

    load_dotenv(BASE_DIR / ".env")
except ImportError:
    pass

import hashlib
import os
import threading
import time
import uuid
from typing import Any, Dict, Optional

import fitz
from flask import Flask, jsonify, render_template, request, send_file

from converter.convertapi_config import convertapi_any_token_set, convertapi_env_label, resolve_convertapi_secret
from converter.excel_builder import build_excel
from converter.pdf_parser import parse_pdf

UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
PREVIEW_DIR = BASE_DIR / "previews"

for d in (UPLOAD_DIR, OUTPUT_DIR, PREVIEW_DIR):
    d.mkdir(parents=True, exist_ok=True)

app = Flask(
    __name__,
    static_folder="static",
    template_folder="templates",
)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024  # 80 MB

# In-memory only: use Gunicorn --workers 1 (see Procfile). Multiple workers → 404 on /convert poll.
_jobs: Dict[str, Dict[str, Any]] = {}
_jobs_lock = threading.Lock()
# Same PDF bytes → same job (avoids double upload when UI fires twice)
_hash_to_file_id: Dict[str, str] = {}


def _job_get(fid: str) -> Dict[str, Any]:
    with _jobs_lock:
        return _jobs.get(fid, {}).copy()


def _job_set(fid: str, **kwargs: Any) -> None:
    with _jobs_lock:
        if fid not in _jobs:
            _jobs[fid] = {}
        _jobs[fid].update(kwargs)


def _parse_bool_arg(val: Optional[str]) -> bool:
    if val is None:
        return False
    return str(val).strip().lower() in ("1", "true", "yes", "on")


def _parse_engine_and_convertapi_options() -> tuple[str, Dict[str, Any]]:
    engine = (request.args.get("engine") or "local").strip().lower()
    if engine not in ("local", "convertapi"):
        engine = "local"
    convertapi_options: Dict[str, Any] = {}
    if engine == "convertapi":
        convertapi_options = {
            "single_sheet": _parse_bool_arg(request.args.get("single_sheet")),
            "include_formatting": _parse_bool_arg(request.args.get("include_formatting")),
        }
    return engine, convertapi_options


def _try_start_conversion_job(
    file_id: str, engine: str, convertapi_options: Dict[str, Any]
) -> None:
    start_thread = False
    with _jobs_lock:
        j = _jobs.get(file_id)
        if j and j.get("status") == "uploaded" and not j.get("converting"):
            _jobs[file_id]["converting"] = True
            _jobs[file_id]["status"] = "converting"
            _jobs[file_id]["engine"] = engine
            if engine == "convertapi":
                _jobs[file_id]["convertapi_options"] = convertapi_options
            else:
                _jobs[file_id].pop("convertapi_options", None)
            start_thread = True
    if start_thread:
        if engine == "convertapi":
            threading.Thread(
                target=_run_convertapi_conversion, args=(file_id,), daemon=True
            ).start()
        else:
            threading.Thread(target=_run_conversion, args=(file_id,), daemon=True).start()


def _render_pdf_previews(pdf_path: Path, fid: str) -> int:
    out_sub = PREVIEW_DIR / fid
    out_sub.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    try:
        n = len(doc)
        zoom = 1.6
        mat = fitz.Matrix(zoom, zoom)
        for i in range(n):
            page = doc[i]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            png_path = out_sub / f"page_{i + 1}.png"
            pix.save(str(png_path))
    finally:
        doc.close()
    return n


def _get_pdf_page_count(pdf_path: Path) -> int:
    """Fast page count without rendering previews."""
    doc = fitz.open(str(pdf_path))
    try:
        return len(doc)
    finally:
        doc.close()


def _render_pdf_previews_limited(pdf_path: Path, fid: str, max_pages: int) -> None:
    """
    Render a limited number of preview PNGs in background.
    This keeps /upload fast on Render, Coolify, or any production host.
    """
    if max_pages <= 0:
        return
    out_sub = PREVIEW_DIR / fid
    out_sub.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    try:
        n = min(len(doc), max_pages)
        zoom = 1.6
        mat = fitz.Matrix(zoom, zoom)
        for i in range(n):
            page = doc[i]
            pix = page.get_pixmap(matrix=mat, alpha=False)
            png_path = out_sub / f"page_{i + 1}.png"
            pix.save(str(png_path))
            _job_set(fid, preview_rendered=i + 1)
    finally:
        doc.close()


def _run_conversion(fid: str) -> None:
    with _jobs_lock:
        job = _jobs.get(fid)
        if not job:
            return
        pdf_path = Path(job["pdf_path"])
        out_path = Path(job["out_path"])

    try:
        _job_set(fid, status="parsing", progress=15, message="Parsing PDF")
        parsed = parse_pdf(str(pdf_path))

        def cb(pct: int, msg: str) -> None:
            # Map 15–95 into builder progress
            scaled = 15 + int(pct * 0.8)
            _job_set(fid, progress=min(95, scaled), message=msg)

        _job_set(fid, status="building", progress=40, message="Building Excel")
        meta = build_excel(parsed, str(out_path), pdf_path=str(pdf_path), progress_cb=cb)
        meta = dict(meta)
        meta["conversion_engine"] = "local"

        _job_set(
            fid,
            status="done",
            progress=100,
            message="Complete",
            excel_meta=meta,
            error=None,
        )
    except Exception as e:
        _job_set(fid, status="error", progress=0, message=str(e), error=str(e))


def _run_convertapi_conversion(fid: str) -> None:
    with _jobs_lock:
        job = _jobs.get(fid)
        if not job:
            return
        pdf_path = Path(job["pdf_path"])
        out_path = Path(job["out_path"])
        opts = dict(job.get("convertapi_options") or {})

    try:
        secret, api_mode = resolve_convertapi_secret()

        single_sheet = bool(opts.get("single_sheet", False))
        include_formatting = bool(opts.get("include_formatting", False))

        _job_set(
            fid,
            status="converting",
            progress=20,
            message=f"ConvertAPI ({api_mode}): uploading PDF",
        )

        from converter.convertapi_client import convert_pdf_to_xlsx_convertapi
        from converter.xlsx_preview import read_xlsx_preview

        api_meta = convert_pdf_to_xlsx_convertapi(
            pdf_path,
            out_path,
            secret,
            single_sheet=single_sheet,
            include_formatting=include_formatting,
        )

        _job_set(fid, progress=75, message="ConvertAPI: building preview")
        headers, rows = read_xlsx_preview(out_path)
        preview_rows = rows[:300]

        excel_meta: Dict[str, Any] = {
            "rows_written": len(rows),
            "cols_written": len(headers),
            "preview_headers": headers,
            "preview_rows": preview_rows,
            "page_count": job.get("preview_pages", 0),
            "used_ocr": False,
            "structured_table_count": 0,
            "sheets": ["Sheet1"],
            "conversion_engine": "convertapi",
            "convertapi_mode": api_mode,
            "convertapi_options": {
                "single_sheet": single_sheet,
                "include_formatting": include_formatting,
            },
            "convertapi": api_meta,
            "visual_match_note": (
                "ConvertAPI extracts tables into spreadsheet cells; it does not reproduce PDF "
                "appearance (logos, fonts, colors, or exact page layout). For layout-styled output "
                "(e.g. TRACES Form 26AS), use the Local engine instead."
            ),
        }

        _job_set(
            fid,
            status="done",
            progress=100,
            message="Complete",
            excel_meta=excel_meta,
            error=None,
        )
    except Exception as e:
        _job_set(fid, status="error", progress=0, message=str(e), error=str(e))


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config", methods=["GET"])
def api_config():
    """Client UI: whether optional ConvertAPI engine is available."""
    return jsonify(
        {
            "convertapi_configured": convertapi_any_token_set(),
            "convertapi_env": convertapi_env_label(),
        }
    )


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400
    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Please upload a PDF file"}), 400

    raw = f.read()
    digest = hashlib.sha256(raw).hexdigest()

    with _jobs_lock:
        existing = _hash_to_file_id.get(digest)
        if existing and existing in _jobs:
            job = _jobs[existing]
            p = Path(job.get("pdf_path", ""))
            # Only merge duplicate POSTs while the job is still "fresh" (not finished)
            if p.is_file() and job.get("status") in ("uploaded", "converting"):
                return jsonify(
                    {
                        "file_id": existing,
                        "page_count": job.get("preview_pages", 0),
                        "original_name": job.get("filename", f.filename),
                        "deduplicated": True,
                    }
                )

    fid = uuid.uuid4().hex
    safe_name = f"{fid}.pdf"
    pdf_path = UPLOAD_DIR / safe_name
    pdf_path.write_bytes(raw)

    out_name = f"{fid}.xlsx"
    out_path = OUTPUT_DIR / out_name

    # IMPORTANT: keep /upload fast everywhere (Render, Coolify, Docker, etc.).
    # Preview PNGs render in a background thread and are capped by MAX_PREVIEW_PAGES.
    try:
        page_count = _get_pdf_page_count(pdf_path)
    except Exception:
        page_count = 0

    try:
        max_preview = int(os.environ.get("MAX_PREVIEW_PAGES", "4"))
    except Exception:
        max_preview = 4
    if max_preview < 0:
        max_preview = 0

    threading.Thread(
        target=_render_pdf_previews_limited,
        args=(pdf_path, fid, max_preview),
        daemon=True,
    ).start()

    with _jobs_lock:
        _hash_to_file_id[digest] = fid
        _jobs[fid] = {
            "pdf_path": str(pdf_path),
            "out_path": str(out_path),
            "filename": f.filename,
            "status": "uploaded",
            "progress": 0,
            "message": "Uploaded",
            "preview_pages": page_count,
            "preview_rendered": 0,
            "error": None,
            "excel_meta": None,
            "sha256": digest,
        }

    return jsonify(
        {
            "file_id": fid,
            "page_count": page_count,
            "original_name": f.filename,
            "deduplicated": False,
        }
    )


def _convert_response_json(file_id: str) -> tuple[Dict[str, Any], int]:
    """Snapshot of job for /convert and /convert_wait (200 unless unknown)."""
    with _jobs_lock:
        job = _jobs.get(file_id)
    if not job:
        return {"error": "Unknown file_id"}, 404

    if job.get("status") == "done":
        return (
            {
                "file_id": file_id,
                "progress": 100,
                "status": "done",
                "message": job.get("message"),
                "excel_meta": job.get("excel_meta"),
            },
            200,
        )

    if job.get("status") == "error":
        return (
            {
                "file_id": file_id,
                "progress": 0,
                "status": "error",
                "message": job.get("error") or job.get("message"),
            },
            200,
        )

    j = _job_get(file_id)
    return (
        {
            "file_id": file_id,
            "progress": j.get("progress", 0),
            "status": j.get("status", "converting"),
            "message": j.get("message", ""),
        },
        200,
    )


@app.route("/convert/<file_id>", methods=["GET"])
def convert(file_id: str):
    engine, convertapi_options = _parse_engine_and_convertapi_options()
    with _jobs_lock:
        job = _jobs.get(file_id)
    if not job:
        return jsonify({"error": "Unknown file_id"}), 404
    _try_start_conversion_job(file_id, engine, convertapi_options)
    body, code = _convert_response_json(file_id)
    return jsonify(body), code


@app.route("/convert_wait/<file_id>", methods=["GET"])
def convert_wait(file_id: str):
    """
    Hold the request until the job finishes, errors, or wait_ms elapses (default 25s).
    Reduces client polling round-trips; keep wait_ms below reverse-proxy timeouts (~100s).
    """
    engine, convertapi_options = _parse_engine_and_convertapi_options()

    with _jobs_lock:
        job = _jobs.get(file_id)
    if not job:
        return jsonify({"error": "Unknown file_id"}), 404

    if job.get("status") in ("done", "error"):
        body, code = _convert_response_json(file_id)
        return jsonify(body), code

    _try_start_conversion_job(file_id, engine, convertapi_options)

    try:
        wait_ms = int(request.args.get("wait_ms", "25000"))
    except (TypeError, ValueError):
        wait_ms = 25000
    wait_ms = max(1000, min(wait_ms, 25000))

    deadline = time.monotonic() + wait_ms / 1000.0
    while time.monotonic() < deadline:
        j = _job_get(file_id)
        st = j.get("status")
        if st == "done":
            body, code = _convert_response_json(file_id)
            return jsonify(body), code
        if st == "error":
            body, code = _convert_response_json(file_id)
            return jsonify(body), code
        time.sleep(0.25)

    body, code = _convert_response_json(file_id)
    return jsonify(body), code


@app.route("/download/<file_id>", methods=["GET"])
def download(file_id: str):
    with _jobs_lock:
        job = _jobs.get(file_id)
    if not job:
        return jsonify({"error": "Unknown file_id"}), 404
    out_path = Path(job["out_path"])
    if not out_path.is_file():
        return jsonify({"error": "File not ready"}), 400
    orig = job.get("filename") or "document.pdf"
    stem = Path(orig).stem
    dl_name = f"{stem}.xlsx"
    return send_file(
        str(out_path),
        as_attachment=True,
        download_name=dl_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/preview/<file_id>", methods=["GET"])
def preview(file_id: str):
    with _jobs_lock:
        job = _jobs.get(file_id)
    if not job:
        return jsonify({"error": "Unknown file_id"}), 404

    prev_dir = PREVIEW_DIR / file_id
    pages = []
    if prev_dir.is_dir():
        pngs = sorted(prev_dir.glob("page_*.png"), key=lambda p: p.name)
        for i, p in enumerate(pngs):
            pages.append(
                {
                    "page": i + 1,
                    "url": f"/preview_page/{file_id}/{i + 1}",
                }
            )

    j = _job_get(file_id)
    excel_preview = (j.get("excel_meta") or {}).get("preview_rows")
    headers = (j.get("excel_meta") or {}).get("preview_headers")

    return jsonify(
        {
            "file_id": file_id,
            "page_count": len(pages) or j.get("preview_pages", 0),
            "pdf_pages": pages,
            "excel_preview": {
                "headers": headers or [],
                "rows": excel_preview or [],
            },
            "status": j.get("status"),
        }
    )


@app.route("/preview_page/<file_id>/<int:page_num>", methods=["GET"])
def preview_page(file_id: str, page_num: int):
    path = PREVIEW_DIR / file_id / f"page_{page_num}.png"
    if not path.is_file():
        return jsonify({"error": "Not found"}), 404
    return send_file(str(path), mimetype="image/png")


if __name__ == "__main__":
    # Default 5001 — macOS often reserves 5000 for AirPlay Receiver
    port = int(os.environ.get("PORT", "5001"))
    app.run(host="0.0.0.0", port=port, debug=True)
