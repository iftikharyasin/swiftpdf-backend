# Optimax PDF API on Render (Docker web service).
# Python version and the full LibreOffice install stay as they were. The move
# to Python 3.12 and a slimmer LibreOffice is still the Backend upgrade step
# (AUDIT.md 6.11 and 6.12). This change adds only the three metric compatible
# font families, which the PDF to Word rebuild needs.
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
#
# The three font packages are new (AUDIT.md 6.12, part of it):
#   fonts-liberation        Liberation Serif, Sans and Mono, which match the
#                           metrics of Times New Roman, Arial and Courier New
#   fonts-crosextra-carlito Carlito, which matches Calibri
#   fonts-crosextra-caladea Caladea, which matches Cambria
# PDF to Word maps PostScript font names onto exactly those four families, so
# with these installed the server lays a page out the way Word will. They also
# improve Word to PDF, which until now had no Microsoft compatible fonts at
# all and reflowed every page (AUDIT.md 3.9). Arabic and Urdu fonts are still
# missing and stay on the Backend upgrade list.
RUN apt-get update && apt-get install -y --no-install-recommends \
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
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    libgl1 \
    libglib2.0-0 \
    tini \
    && fc-cache -f \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY server.py .
COPY docx_repair.py .

# tini runs first so killed jobs never leave zombie processes behind.
ENTRYPOINT ["/usr/bin/tini", "--"]

# One API process. Heavy work runs in child processes started by server.py.
# --limit-concurrency only trips under a flood; normal traffic waits in the
# job queue inside server.py and gets a clear "busy" reply.
CMD ["sh", "-c", "exec uvicorn server:app --host 0.0.0.0 --port ${PORT:-10000} --limit-concurrency 64 --timeout-keep-alive 5"]
