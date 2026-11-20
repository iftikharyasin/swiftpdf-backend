# Optimax PDF API on Render (Docker web service).
# Python version and the full LibreOffice install stay as they were. The move
# to Python 3.12 and a slimmer LibreOffice is still the Backend upgrade step
# (AUDIT.md 6.11 and 6.12).
#
# 9 November 2026: the three metric compatible font families, which the PDF to
# Word rebuild needs.
# 20 November 2026: fonts-noto-core, so documents in Arabic and Urdu convert
# instead of coming back as empty boxes (AUDIT.md 3.9 and 6.12).
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
# The fonts, and why each one is here (AUDIT.md 6.12):
#   fonts-liberation        Liberation Serif, Sans and Mono, which match the
#                           metrics of Times New Roman, Arial and Courier New
#   fonts-crosextra-carlito Carlito, which matches Calibri
#   fonts-crosextra-caladea Caladea, which matches Cambria
#   fonts-noto-core         Noto, covering Arabic and most other scripts. It
#                           carries Noto Naskh Arabic, Noto Kufi Arabic and
#                           Noto Nastaliq Urdu, Nastaliq being the style Urdu
#                           is actually written in, so this one package serves
#                           both scripts. There is no separate Debian package
#                           for Nastaliq; asking for one fails the build.
#
# PDF to Word maps PostScript font names onto the first four families, so with
# these installed the server lays a page out the way Word will. Word to PDF
# gets the same benefit: until 9 November it had no Microsoft compatible fonts
# at all and reflowed every page (AUDIT.md 3.9).
#
# Noto enlarges the image and so lengthens a cold start. Watch the first
# deploy after adding it. fonts-noto-core is already the trimmed package;
# fonts-noto, the full set, is far larger and is not worth it here.
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
    fonts-noto-core \
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
COPY pdf_prepare.py .

# tini runs first so killed jobs never leave zombie processes behind.
ENTRYPOINT ["/usr/bin/tini", "--"]

# One API process. Heavy work runs in child processes started by server.py.
# --limit-concurrency only trips under a flood; normal traffic waits in the
# job queue inside server.py and gets a clear "busy" reply.
CMD ["sh", "-c", "exec uvicorn server:app --host 0.0.0.0 --port ${PORT:-10000} --limit-concurrency 64 --timeout-keep-alive 5"]
