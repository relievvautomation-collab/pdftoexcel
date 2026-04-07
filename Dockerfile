# PDF → Excel (Flask) — Coolify, Fly.io, or any Docker host
FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=5001

WORKDIR /app

# System deps: Camelot (Ghostscript), pdfplumber/tabula (Poppler), Tabula (Java), OCR (Tesseract)
RUN apt-get update && apt-get install -y --no-install-recommends \
    ghostscript \
    poppler-utils \
    default-jre-headless \
    tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY . .

RUN mkdir -p uploads outputs previews

EXPOSE 5001

CMD gunicorn app:app --bind 0.0.0.0:${PORT} --workers 2 --threads 4 --timeout 300 --access-logfile - --error-logfile -
