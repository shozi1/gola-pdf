# ==============================================================================
# Gola PDF - High Performance Production Container (FastAPI + PyMuPDF + OCR)
# Optimized for Hostinger VPS (KVM 2 / KVM 4) or any Docker-enabled Linux Host
# ==============================================================================

FROM python:3.11-slim-bookworm

# Prevent Python from writing .pyc files & enable unbuffered standard I/O
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    PORT=8080

# Install system dependencies: Tesseract OCR, Ghostscript, C build tools
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-fra \
    tesseract-ocr-deu \
    tesseract-ocr-spa \
    ghostscript \
    curl \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python requirements
COPY requirements.txt /app/
RUN pip install --no-cache-dir -U pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt && \
    pip install --no-cache-dir gunicorn uvicorn

# Copy application source code
COPY . /app/

# Create ephemeral storage directories with appropriate permissions
RUN mkdir -p /app/ephemeral_gcs_storage && \
    chmod -R 777 /app/ephemeral_gcs_storage

EXPOSE 8080

# Run with Gunicorn process manager and Uvicorn workers
CMD ["sh", "-c", "gunicorn -w 2 -k uvicorn.workers.UvicornWorker app:app --bind 0.0.0.0:${PORT:-8080} --timeout 120"]
