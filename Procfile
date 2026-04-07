# Single worker: job state lives in process memory (_jobs). Multiple workers cause 404 on /convert polling.
web: gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --worker-class gthread --threads 8 --timeout 300 --access-logfile - --error-logfile -
