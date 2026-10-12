FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=5000
ENV PYTHONUNBUFFERED=1
ENV DATA_DIR=/tmp/alrawasi-data
CMD ["sh", "-c", "python prepare_runtime_db.py && exec gunicorn --bind 0.0.0.0:${PORT:-5000} --workers 1 --threads 1 --timeout 120 free_seed_bootstrap:flask_app"]
