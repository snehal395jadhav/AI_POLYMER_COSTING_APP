# syntax=docker/dockerfile:1.6
# ---------- Costing Desk — production image ----------
FROM python:3.12-slim AS base

# Prevent Python from writing .pyc files and force stdout flushing
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Install runtime OS packages only (curl is for healthcheck)
RUN apt-get update \
 && apt-get install -y --no-install-recommends curl \
 && rm -rf /var/lib/apt/lists/*

# Create non-root user
RUN groupadd -r app && useradd -r -g app -d /app -s /sbin/nologin app

WORKDIR /app

# Install Python deps first (better layer caching)
COPY requirements.txt ./
RUN pip install -r requirements.txt

# Copy application source
COPY --chown=app:app . .

# Writable data directory (SQLite lives here; mount a volume in compose)
RUN mkdir -p /app/data && chown -R app:app /app

USER app

EXPOSE 5000

# Simple container healthcheck hits the login page
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -fsS http://127.0.0.1:5000/login || exit 1

# Production WSGI server: 2 workers × 4 threads is a safe default for SQLite
CMD ["gunicorn", \
     "--bind=0.0.0.0:5000", \
     "--workers=2", \
     "--threads=4", \
     "--access-logfile=-", \
     "--error-logfile=-", \
     "app:app"]
