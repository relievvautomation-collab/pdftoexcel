# PDF to Excel Converter

Flask app: upload a PDF, convert to `.xlsx` (local layout-aware pipeline or optional **ConvertAPI** cloud).

## Requirements

- Python 3.12+ recommended
- For **full local conversion** (Camelot / Tabula): **Ghostscript**, **Poppler**, **Java** (Tabula), optional **Tesseract** for OCR — see [`Dockerfile`](Dockerfile) for the Debian package list.

## Local setup

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```
 
Copy [`env.example`](env.example) to `.env` and fill secrets (never commit `.env`).

```bash
export PORT=5001
python app.py
```

Open `http://localhost:5001`.

### Production-style run (local)

```bash
gunicorn app:app --bind 0.0.0.0:5001 --workers 2 --threads 4 --timeout 300
```

## Environment variables

| Variable | Purpose |
|----------|---------|
| `PORT` | HTTP port (Render/Coolify set this automatically) |
| `CONVERTAPI_SECRET` / `CONVERTAPI_SECRET_SANDBOX` | ConvertAPI tokens |
| `CONVERTAPI_ENV` | `production` or `sandbox` |
| `CONVERTAPI_BASE_URI` | Optional regional API base URL |

See [`env.example`](env.example).

## Deploy on Render

1. Push this repo to GitHub (no `.env`, no `uploads/` / `outputs/` / `previews/` — see [`.gitignore`](.gitignore)).
2. New **Web Service** → connect repo.
3. **Build command:** `pip install -r requirements.txt`
4. **Start command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 300`  
   (or rely on [`Procfile`](Procfile) if Render detects it.)
5. **Environment:** add the same variables as in `env.example` in the Render dashboard.
6. Optional: [`runtime.txt`](runtime.txt) pins Python for Render.

**Note:** Ephemeral disk — uploaded files may be lost on restart. For production, add persistent disk or S3.

## Deploy on Coolify (Docker)

1. Push repo to GitHub.
2. New resource → **Dockerfile** build from this repo.
3. Set environment variables in Coolify (same as `env.example`).
4. Map HTTP port to container port **`5001`** (or set `PORT` to the port Coolify injects and ensure the app binds to it).

The [`Dockerfile`](Dockerfile) installs system libraries for Camelot/Tabula/OCR inside the image.

## Repository contents

- [`app.py`](app.py) — Flask entrypoint (`app` object for Gunicorn).
- [`requirements.txt`](requirements.txt) — Python dependencies.
- [`Procfile`](Procfile) — process type for Render/Heroku-style hosts.
- [`Dockerfile`](Dockerfile) — container image for Coolify / Docker.
- [`env.example`](env.example) — template for secrets (copy to `.env` locally only).

## License

Add your license if needed.
