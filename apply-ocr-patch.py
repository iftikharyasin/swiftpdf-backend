#!/usr/bin/env python3
"""Patch server.py for the OCR PDF rebuild (21 November 2026).

Run it from the folder that holds server.py:

    python apply-ocr-patch.py server.py

It refuses to do anything unless it finds the exact text it expects, it
compiles the result before writing, and it keeps a copy as server.py.bak.
Running it twice is safe: the second run says the patch is already applied
and changes nothing.

What it changes, and why.

1. Tesseract is asked what language packs it actually has, once, at the first
   request that needs the list. Until now the server took any string that
   looked like a language code and found out it was wrong only when OCRmyPDF
   failed, which closes AUDIT.md 6.14. The list also goes out on /health, so
   the page can only ever offer languages the server will accept.

2. task_ocr stops forcing force_ocr on every job (also AUDIT.md 6.14). The
   page now chooses:
       mode=skip   skip_text, leave pages that already hold text alone
       mode=redo   redo_ocr, replace a text layer an earlier scanner added
       mode=force  force_ocr, exactly what this server has always done
   No mode field at all still means force, so a page that has not been
   updated behaves precisely as before.

3. Image clean up before reading: rotate_pages, deskew, remove_background and
   clean. OCRmyPDF refuses all four alongside redo_ocr, because that mode
   compares the page with what is already drawn on it, so they are dropped in
   that mode here as well as being greyed out on the page.

4. A page range, so a long file can be read in parts.

5. output=pdfa or output=pdf. OCRmyPDF's own default is pdfa and nothing was
   passed before, so pdfa stays the default here and nothing changes for a
   request that does not ask.

6. want_text=true writes the recognised text beside the PDF and sends both
   back as a ZIP.
"""

import os
import shutil
import sys

# ---------------------------------------------------------------------------
# 1. Language discovery, the mode and output lists
# ---------------------------------------------------------------------------

ANCHOR_PATTERN = 'OCR_LANGUAGE_PATTERN = re.compile(r"[a-z_]{3,10}(\\+[a-z_]{3,10}){0,4}")'

LANGUAGES_BLOCK = '''

# OCR PDF (AUDIT.md 6.14). Modes and output types the page may ask for.
# An unknown value falls back to what this server did before the patch, so a
# page that has not been updated is unaffected.
OCR_MODES = ("skip", "redo", "force")
OCR_OUTPUTS = ("pdfa", "pdf")

# Tesseract ships these two as data files, but neither is a language: osd is
# orientation and script detection, equ is the equation model.
OCR_NOT_A_LANGUAGE = frozenset({"osd", "equ"})

# At most this many languages in one pass. Each one Tesseract loads costs
# memory, and the free tier has 512 MB for everything.
OCR_MAX_LANGUAGES = 3

_ocr_languages_cache = None


def ocr_languages() -> tuple:
    """The language packs Tesseract actually has, asked of Tesseract itself.

    Read once and kept. Guessing from a list in the code would drift from the
    image the moment the Dockerfile changed, which is exactly the fault this
    closes: a language the page offered and the server did not have failed
    only after a full upload.
    """
    global _ocr_languages_cache
    if _ocr_languages_cache is None:
        found = []
        try:
            answer = subprocess.run(
                ["tesseract", "--list-langs"],
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
            # Tesseract prints a header line with spaces in it, then one code
            # per line. Some builds print the whole thing on stderr.
            for line in ((answer.stdout or "") + "\\n" + (answer.stderr or "")).splitlines():
                code = line.strip()
                if code and " " not in code and ":" not in code and code not in OCR_NOT_A_LANGUAGE:
                    found.append(code)
        except Exception:  # noqa: BLE001 - a missing binary must not stop the server
            found = []
        _ocr_languages_cache = tuple(sorted(set(found))) or ("eng",)
        # `log` belongs to the API process, which is defined further down this
        # file. The worker child exits before that line is ever reached, so
        # the logger is asked for rather than assumed.
        logger = globals().get("log")
        if logger is not None:
            logger.info("ocr languages=%d", len(_ocr_languages_cache))
    return _ocr_languages_cache
'''

# ---------------------------------------------------------------------------
# 2. The worker
# ---------------------------------------------------------------------------

OLD_TASK_OCR = '''def task_ocr(job_dir: str, params: dict) -> dict:
    src = os.path.join(job_dir, "input.pdf")
    out = os.path.join(job_dir, "output.pdf")

    pages = _check_pdf(src, 403, reject_any_encryption=True)
    _check_page_cap(pages, int(params.get("max_pages") or 0))

    import ocrmypdf

    try:
        ocrmypdf.ocr(
            src,
            out,
            language=params.get("language", "eng"),
            force_ocr=True,
            optimize=0,
            jobs=1,
            skip_big=10.0,
        )
    except MemoryError:
        raise
    except Exception as exc:
        raise _ocrmypdf_error_to_api(exc, 403)

    _require_output(out)
    return {}'''

NEW_TASK_OCR = '''def task_ocr(job_dir: str, params: dict) -> dict:
    src = os.path.join(job_dir, "input.pdf")
    out = os.path.join(job_dir, "output.pdf")
    sidecar = os.path.join(job_dir, "output.txt")

    pages = _check_pdf(src, 403, reject_any_encryption=True)
    _check_page_cap(pages, int(params.get("max_pages") or 0))

    import ocrmypdf

    # No mode is the behaviour this server has always had, so an older page
    # gets exactly the file it got yesterday.
    mode = params.get("mode") or "force"

    options = {
        "language": params.get("language", "eng"),
        "optimize": 0,
        "jobs": 1,
        "skip_big": 10.0,
        # OCRmyPDF's own default is already pdfa. Saying so makes the output
        # a decision rather than an accident of the library's defaults.
        "output_type": params.get("output") or "pdfa",
    }
    if mode == "skip":
        # Pages that already hold real text are copied through untouched, so
        # their text stays vector and the file does not grow.
        options["skip_text"] = True
    elif mode == "redo":
        options["redo_ocr"] = True
    else:
        options["force_ocr"] = True

    # Image clean up. OCRmyPDF refuses every one of these alongside redo_ocr,
    # because that mode works by comparing the page with what is drawn on it.
    # Dropping them here means a page that asks anyway gets a result instead
    # of a failed job.
    if mode != "redo":
        if params.get("rotate"):
            options["rotate_pages"] = True
        if params.get("deskew"):
            options["deskew"] = True
        if params.get("background"):
            options["remove_background"] = True
        if params.get("clean"):
            options["clean"] = True

    page_range = params.get("pages") or ""
    if page_range:
        options["pages"] = page_range

    want_text = bool(params.get("want_text"))
    if want_text:
        options["sidecar"] = sidecar

    try:
        ocrmypdf.ocr(src, out, **options)
    except MemoryError:
        raise
    except Exception as exc:
        raise _ocrmypdf_error_to_api(exc, 403)

    _require_output(out)

    if want_text and os.path.exists(sidecar):
        import zipfile

        bundle = os.path.join(job_dir, "output.zip")
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.write(out, "searchable.pdf")
            archive.write(sidecar, "text.txt")
        _require_output(bundle)
        return {"zipped": True}

    return {}'''

# ---------------------------------------------------------------------------
# 3. The endpoint
# ---------------------------------------------------------------------------

OLD_SIZE_LIMIT = '''    "/api/ocr": (10, 180),'''

NEW_SIZE_LIMIT = '''    # 15 MB since the OCR rebuild of 21 November, to match the page and
    # iLovePDF. Safe because the default mode no longer rasterises pages that
    # already hold text, which was the memory cost on every job.
    "/api/ocr": (15, 180),'''

OLD_ENDPOINT = '''@app.post("/api/ocr")
def ocr_pdf(file: UploadFile = File(...), language: str = Form("eng")):
    with Job("/api/ocr", file) as job:
        if not OCR_LANGUAGE_PATTERN.fullmatch(language or ""):
            raise ApiError(400, "bad_language")
        job.start(".pdf")
        job.run_worker("ocr", {"language": language, "max_pages": MAX_PAGES_OCR})
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_Searchable_{job.base}.pdf")'''

NEW_ENDPOINT = '''@app.post("/api/ocr")
def ocr_pdf(
    file: UploadFile = File(...),
    language: str = Form("eng"),
    mode: str = Form(""),
    rotate: bool = Form(False),
    deskew: bool = Form(False),
    background: bool = Form(False),
    clean: bool = Form(False),
    pages: str = Form(""),
    output: str = Form(""),
    want_text: bool = Form(False),
):
    """OCR PDF.

    Every field but `file` and `language` is new and optional. A request that
    sends neither gets the same job, the same settings and the same file name
    the old page received.
    """
    with Job("/api/ocr", file) as job:
        if not OCR_LANGUAGE_PATTERN.fullmatch(language or ""):
            raise ApiError(400, "bad_language")

        # Checked against the packs this image actually holds, rather than
        # against a shape (AUDIT.md 6.14). A language we do not have is
        # refused here, before the upload is turned into a job.
        wanted = [code for code in (language or "eng").split("+") if code]
        have = set(ocr_languages())
        if not wanted or len(wanted) > OCR_MAX_LANGUAGES or any(code not in have for code in wanted):
            raise ApiError(400, "bad_language")

        chosen_mode = (mode or "").strip().lower()
        chosen_mode = chosen_mode if chosen_mode in OCR_MODES else "force"

        chosen_output = (output or "").strip().lower()
        chosen_output = chosen_output if chosen_output in OCR_OUTPUTS else "pdfa"

        page_range = (pages or "").strip()
        if page_range and not PAGE_RANGE_PATTERN.fullmatch(page_range):
            raise ApiError(400, "bad_pages")

        job.start(".pdf")
        result = job.run_worker(
            "ocr",
            {
                "language": language,
                "max_pages": MAX_PAGES_OCR,
                "mode": chosen_mode,
                "rotate": bool(rotate),
                "deskew": bool(deskew),
                "background": bool(background),
                "clean": bool(clean),
                "pages": page_range,
                "output": chosen_output,
                "want_text": bool(want_text),
            },
        )
        if result.get("zipped"):
            return job.file_response(
                "output.zip", "application/zip", f"OptimaxPDF_Searchable_{job.base}.zip"
            )
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_Searchable_{job.base}.pdf")'''

# ---------------------------------------------------------------------------
# 4. /health reports the installed languages
# ---------------------------------------------------------------------------

HEALTH_DEF = "async def health():"
HEALTH_KEY = '"ocr_languages": list(ocr_languages()), '


def patch_health(source: str) -> str:
    """Add ocr_languages to whatever dict /health already returns.

    This one is done by hand rather than by exact text, because /health has
    picked up a field or two since this patch was written and will pick up
    more. The transformation is narrow: find the first `return {` inside
    `async def health()` and put one key straight after the brace. It works
    whether the dict is written on one line or several, and it keeps every
    field already there.
    """
    if source.count(HEALTH_DEF) != 1:
        raise LookupError(
            f"expected exactly one '{HEALTH_DEF}' and found {source.count(HEALTH_DEF)}"
        )
    start = source.index(HEALTH_DEF)
    brace = source.find("return {", start)
    if brace == -1 or brace - start > 1200:
        raise LookupError("could not find a 'return {' inside async def health()")
    cut = brace + len("return {")
    # No trailing space when the dict carries on to the next line.
    key = HEALTH_KEY.rstrip() if source[cut:cut + 1] == "\n" else HEALTH_KEY
    return source[:cut] + key + source[cut:]


def fail(message: str) -> None:
    print(f"Nothing was changed. {message}")
    sys.exit(1)


def main() -> None:
    path = sys.argv[1] if len(sys.argv) > 1 else "server.py"
    if not os.path.exists(path):
        fail(f"{path} is not here. Run this from the folder that holds server.py.")

    with open(path, encoding="utf-8") as handle:
        source = handle.read()

    if "def ocr_languages()" in source:
        print(f"{path} already has the OCR patch. Nothing to do.")
        return

    if "import subprocess" not in source:
        fail("server.py does not import subprocess, so this is not the file this patch was written for.")

    missing = []
    if ANCHOR_PATTERN not in source:
        missing.append("the OCR_LANGUAGE_PATTERN line")
    if OLD_TASK_OCR not in source:
        missing.append("task_ocr as it was before this patch")
    if OLD_ENDPOINT not in source:
        missing.append("the /api/ocr endpoint as it was before this patch")
    if OLD_SIZE_LIMIT not in source:
        missing.append("the /api/ocr size limit as it was before this patch")
    if missing:
        fail(
            "This does not look like the server.py this patch was written for.\n"
            "  Could not find: " + ", ".join(missing) + "\n"
            "  Send me the current server.py and I will rewrite the patch against it."
        )

    patched = source.replace(ANCHOR_PATTERN, ANCHOR_PATTERN + LANGUAGES_BLOCK, 1)
    patched = patched.replace(OLD_TASK_OCR, NEW_TASK_OCR, 1)
    patched = patched.replace(OLD_SIZE_LIMIT, NEW_SIZE_LIMIT, 1)
    patched = patched.replace(OLD_ENDPOINT, NEW_ENDPOINT, 1)
    try:
        patched = patch_health(patched)
    except LookupError as err:
        fail(f"The /health endpoint is not the shape this patch can edit ({err}).")

    try:
        compile(patched, path, "exec")
    except SyntaxError as err:
        fail(f"The patched file would not compile ({err}). server.py is untouched.")

    shutil.copyfile(path, path + ".bak")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(patched)

    print(f"Patched {path}. The file as it was is saved as {path}.bak.")
    print("Changed: language discovery from Tesseract, reading modes, image clean up,")
    print("page range, PDF/A or standard PDF, a text file beside the PDF, /health,")
    print("and the /api/ocr size limit, now 15 MB to match the page.")
    print()
    print("Your /health now reads:")
    start = patched.index(HEALTH_DEF)
    for line in patched[start:].splitlines()[:12]:
        print("    " + line)
        if line.strip().endswith("}") and "return" not in line:
            break
    print()
    print("Please paste those lines back to me so the change is on record.")


if __name__ == "__main__":
    main()
