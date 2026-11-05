# Optimax PDF API on Render (Docker web service).
# Python version and the full LibreOffice install stay as they were in this
# patch. The move to Python 3.12, a slimmer LibreOffice and extra fonts is the
# Backend upgrade step (AUDIT.md 6.11 and 6.12).
FROM python:3.10-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    OMP_THREAD_LIMIT=1 \
    OPTIMAX_JOBS_ROOT=/tmp/optimax-jobs \
    OPTIMAX_LO_PROFILES_ROOT=/tmp/optimax-lo

# Tesseract OCR with the nine languages the OCR page offers, Ghostscript,
# LibreOffice, the graphics libraries pdf2docx needs, and tini, a tiny init
# process that cleans up any stopped child processes after a timeout.
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-spa \
    tesseract-ocr-fra \
    tesseract-ocr-deu \
    tesseract-ocr-ita \
    tesseract-ocr-por \
    tesseract-ocr-nld \
    tesseract-ocr-ara \
    tesseract-ocr-chi-sim \
    ghostscript \
    libreoffice \
    libgl1 \
    libglib2.0-0 \
    tini \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY server.py .

# tini runs first so killed jobs never leave zombie processes behind.
ENTRYPOINT ["/usr/bin/tini", "--"]

# One API process. Heavy work runs in child processes started by server.py.
# --limit-concurrency only trips under a flood; normal traffic waits in the
# job queue inside server.py and gets a clear "busy" reply.
CMD ["sh", "-c", "exec uvicorn server:app --host 0.0.0.0 --port ${PORT:-10000} --limit-concurrency 64 --timeout-keep-alive 5"]
