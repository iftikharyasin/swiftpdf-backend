# Optimax PDF API on Render (Docker web service).
# Python version and the full LibreOffice install stay as they were. The move
# to Python 3.12 and a slimmer LibreOffice is still the Backend upgrade step
# (AUDIT.md 6.11 and 6.12).
#
# 9 November 2026: the three metric compatible font families, which the PDF to
# Word rebuild needs.
# 20 November 2026: fonts-noto-core, so documents in Arabic and Urdu convert
# instead of coming back as empty boxes (AUDIT.md 3.9 and 6.12), plus a
# fontconfig rule so Urdu comes out in Nastaliq rather than Kufi.
# 21 November 2026: unpaper, which is what OCRmyPDF's clean option runs, and
# every remaining Tesseract language pack (AUDIT.md 6.14 and 7.5).
FROM python:3.10-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    OMP_THREAD_LIMIT=1 \
    OPTIMAX_JOBS_ROOT=/tmp/optimax-jobs \
    OPTIMAX_LO_PROFILES_ROOT=/tmp/optimax-lo

# Tesseract OCR with the nine languages this server had from the start,
# Ghostscript, LibreOffice, the graphics libraries pdf2docx needs, and tini, a
# tiny init process that cleans up any stopped child processes after a timeout.
# The other 110 language packs are a block of their own further down, so this
# layer stays cached when that list changes.
#
# New on 21 November:
#   unpaper            what OCRmyPDF's --clean runs to take scanning marks off
#                      a page before reading it. About 1 MB. Without it the
#                      clean option fails the job rather than being ignored.
#   tesseract-ocr-osd  orientation and script detection, which is what
#                      --rotate-pages uses to work out which way a page is up.
#                      tesseract-ocr already depends on it; naming it here
#                      records why it has to stay.
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
    tesseract-ocr-osd \
    ghostscript \
    unpaper \
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

# ---------------------------------------------------------------------------
# The rest of the Tesseract language packs (21 November 2026)
#
# Every language Debian builds from the tesseract-lang source that is not
# already installed above, and none of the tesseract-ocr-script-* packages.
# The script packages are models organised by writing system rather than by
# language, the Han, Japanese and Korean ones are tens of MB each, and the
# individual language packs here already cover those languages.
# tesseract-ocr-all would have pulled all of them in, which is why this is an
# explicit list.
#
# Most packs are between half a megabyte and two megabytes. The build log
# prints the exact figure before it installs them: read the line that begins
# "After this operation" and put that number in CHANGELOG.md.
#
# This costs disk and cold start, not memory. Tesseract loads a language file
# only when a job asks for it, so a hundred installed packs make an English
# job no heavier than it was with nine.
#
# The API does not keep its own copy of this list. server.py asks Tesseract
# with `tesseract --list-langs` and refuses any language that is not in the
# answer, so this file and the API cannot drift apart.
#
# The -vert packs (chi-sim-vert, chi-tra-vert, jpn-vert, kor-vert) are for
# vertically set East Asian text and are deliberately left out. Add them here
# if anyone asks for vertical Japanese or Chinese.
#
# A name that does not exist stops the build with "Unable to locate package".
# Every name below was checked against the Debian tesseract-lang package list
# on 21 November 2026.
# ---------------------------------------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr-afr tesseract-ocr-amh tesseract-ocr-asm tesseract-ocr-aze \
    tesseract-ocr-aze-cyrl tesseract-ocr-bel tesseract-ocr-ben tesseract-ocr-bod \
    tesseract-ocr-bos tesseract-ocr-bre tesseract-ocr-bul tesseract-ocr-cat \
    tesseract-ocr-ceb tesseract-ocr-ces tesseract-ocr-chi-tra tesseract-ocr-chr \
    tesseract-ocr-cos tesseract-ocr-cym tesseract-ocr-dan tesseract-ocr-div \
    tesseract-ocr-dzo tesseract-ocr-ell tesseract-ocr-enm tesseract-ocr-epo \
    tesseract-ocr-est tesseract-ocr-eus tesseract-ocr-fao tesseract-ocr-fas \
    tesseract-ocr-fil tesseract-ocr-fin tesseract-ocr-frk tesseract-ocr-frm \
    tesseract-ocr-fry tesseract-ocr-gla tesseract-ocr-gle tesseract-ocr-glg \
    tesseract-ocr-grc tesseract-ocr-guj tesseract-ocr-hat tesseract-ocr-heb \
    tesseract-ocr-hin tesseract-ocr-hrv tesseract-ocr-hun tesseract-ocr-hye \
    tesseract-ocr-iku tesseract-ocr-ind tesseract-ocr-isl tesseract-ocr-ita-old \
    tesseract-ocr-jav tesseract-ocr-jpn tesseract-ocr-kan tesseract-ocr-kat \
    tesseract-ocr-kat-old tesseract-ocr-kaz tesseract-ocr-khm tesseract-ocr-kir \
    tesseract-ocr-kmr tesseract-ocr-kor tesseract-ocr-lao tesseract-ocr-lat \
    tesseract-ocr-lav tesseract-ocr-lit tesseract-ocr-ltz tesseract-ocr-mal \
    tesseract-ocr-mar tesseract-ocr-mkd tesseract-ocr-mlt tesseract-ocr-mon \
    tesseract-ocr-mri tesseract-ocr-msa tesseract-ocr-mya tesseract-ocr-nep \
    tesseract-ocr-nor tesseract-ocr-oci tesseract-ocr-ori tesseract-ocr-pan \
    tesseract-ocr-pol tesseract-ocr-pus tesseract-ocr-que tesseract-ocr-ron \
    tesseract-ocr-rus tesseract-ocr-san tesseract-ocr-sin tesseract-ocr-slk \
    tesseract-ocr-slv tesseract-ocr-snd tesseract-ocr-spa-old tesseract-ocr-sqi \
    tesseract-ocr-srp tesseract-ocr-srp-latn tesseract-ocr-sun tesseract-ocr-swa \
    tesseract-ocr-swe tesseract-ocr-syr tesseract-ocr-tam tesseract-ocr-tat \
    tesseract-ocr-tel tesseract-ocr-tgk tesseract-ocr-tha tesseract-ocr-tir \
    tesseract-ocr-ton tesseract-ocr-tur tesseract-ocr-uig tesseract-ocr-ukr \
    tesseract-ocr-urd tesseract-ocr-uzb tesseract-ocr-uzb-cyrl tesseract-ocr-vie \
    tesseract-ocr-yid tesseract-ocr-yor \
    && rm -rf /var/lib/apt/lists/*

# Urdu comes out in Kufi without this. fonts-noto-core carries Nastaliq, Naskh
# and Kufi, all covering the Arabic script, so fontconfig was free to answer an
# Urdu request with the squared display face. 99 so it runs last and wins.
COPY urdu-nastaliq.conf /etc/fonts/conf.d/99-urdu-nastaliq.conf
RUN fc-cache -f

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
