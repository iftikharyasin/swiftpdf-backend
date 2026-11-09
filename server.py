"""Optimax PDF API (server.py).

One FastAPI app for every server tool on optimaxpdf.com.

How it stays safe on a small server:
- Uploaded files get fixed internal names inside a private job folder. The
  user's file name is only cleaned and used for the download name.
- Endpoints are plain `def` functions, so FastAPI runs them in worker threads
  and the health check never waits behind a conversion.
- Only MAX_JOBS jobs run at once (1 on the free tier). A few more may wait in a
  short queue; anyone after that gets a clear "busy" reply.
- Every heavy step runs in a short lived child process with a hard timeout.
  A timeout or an out of memory kill stops that one job, not the server.
- LibreOffice uses its own profile per job slot, and a broken profile is reset.
- Request size is checked from Content-Length before the upload is stored, and
  again while the body streams in.
- Only our own website may call the API from a browser, and each visitor has a
  simple rate limit.
- Error replies never contain raw exception text. Logs never contain file names.

Start the API:   uvicorn server:app --host 0.0.0.0 --port 10000
Child process:   python server.py --worker <task> <job_dir>   (started by the API)

Every error reply has this JSON shape:
    {"code": "...", "message": "...", "detail": "..."}
`detail` repeats `message` because the current pages read `detail`.
"""

from __future__ import annotations

import json
import logging
import os
import queue
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from collections import deque


# ---------------------------------------------------------------------------
# Settings. Each one can be changed with an environment variable on Render,
# so moving to a paid tier needs no code change.
# ---------------------------------------------------------------------------

def _env_int(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


def _env_float(name: str, default: float) -> float:
    try:
        return float(os.environ.get(name, default))
    except (TypeError, ValueError):
        return default


MB = 1024 * 1024

# Browser origins that may call the API. www stays allowed until the www to
# bare host redirect is live, so no visitor is locked out during the move.
ALLOWED_ORIGINS = [
    origin.strip().rstrip("/")
    for origin in os.environ.get(
        "OPTIMAX_ALLOWED_ORIGINS",
        "https://optimaxpdf.com,https://www.optimaxpdf.com",
    ).split(",")
    if origin.strip()
]

JOBS_ROOT = os.environ.get("OPTIMAX_JOBS_ROOT", "/tmp/optimax-jobs")
LO_PROFILES_ROOT = os.environ.get("OPTIMAX_LO_PROFILES_ROOT", "/tmp/optimax-lo")
LIBREOFFICE_BIN = os.environ.get("OPTIMAX_LIBREOFFICE_BIN", "libreoffice")

MAX_JOBS = max(1, _env_int("OPTIMAX_MAX_JOBS", 1))
MAX_WAITING = max(0, _env_int("OPTIMAX_MAX_WAITING", 4))
QUEUE_WAIT_SECONDS = max(1, _env_int("OPTIMAX_QUEUE_WAIT_SECONDS", 90))
BUSY_RETRY_AFTER_SECONDS = 10

TIMEOUT_FACTOR = max(0.1, _env_float("OPTIMAX_TIMEOUT_FACTOR", 1.0))

RATE_LIMIT_JOBS = _env_int("OPTIMAX_RATE_LIMIT_JOBS", 30)
RATE_LIMIT_WINDOW_SECONDS = max(1, _env_int("OPTIMAX_RATE_LIMIT_WINDOW_SECONDS", 600))

STALE_JOB_SECONDS = max(60, _env_int("OPTIMAX_STALE_JOB_SECONDS", 900))

# Page caps give a fast, clear answer instead of a long wait for a timeout.
# 0 turns a cap off.
MAX_PAGES_OCR = _env_int("OPTIMAX_MAX_PAGES_OCR", 50)
MAX_PAGES_PDF_TO_WORD = _env_int("OPTIMAX_MAX_PAGES_PDF_TO_WORD", 150)
MAX_PAGES_PDF_TO_EXCEL_OCR = _env_int("OPTIMAX_MAX_PAGES_PDF_TO_EXCEL_OCR", 50)

# Size limit (MB) and timeout (seconds) per endpoint. The size limits must
# match the numbers the pages show (TOOLS.md section 3). Timeouts follow the
# targets in TOOLS.md section 3.1.
TOOL_LIMITS = {
    "/api/convert": (15, 180),
    "/api/word-to-pdf": (20, 120),
    "/api/compress": (50, 120),
    "/api/protect": (30, 60),
    "/api/unlock": (30, 60),
    "/api/ocr": (10, 180),
    "/api/excel-to-pdf": (20, 120),
    "/api/ppt-to-pdf": (20, 120),
    "/api/pdf-to-excel": (20, 180),
}

# Room for multipart headers and form fields around the file itself.
MULTIPART_SLACK_BYTES = 1 * MB

WORKER_FILE = os.path.abspath(__file__)

OCR_LANGUAGE_PATTERN = re.compile(r"[a-z_]{3,10}(\+[a-z_]{3,10}){0,4}")


# ---------------------------------------------------------------------------
# Error codes and the messages users see.
# Some pages still search the message text: OCR looks for "password protected"
# and "unlock", Protect looks for "encrypted", "password" and "protected".
# Only encryption messages may contain those words.
# ---------------------------------------------------------------------------

MESSAGES = {
    "encrypted": "This file is password protected. Please unlock it first.",
    "already_encrypted": "This file is already encrypted. Please unlock it first.",
    "wrong_password": "Incorrect password. Please verify and try again.",
    "password_needed": "This PDF needs a password to open. Enter the password and try again.",
    "corrupt": "We could not read this file. It may be damaged. Try saving it again from the app that made it.",
    "convert_failed": "We could not convert this file. It may be damaged, or saved in a format we cannot read.",
    "empty_file": "This file is empty. Please choose another file.",
    "no_tables": "No tabular data detected in this PDF. If this is a scanned document, please enable the OCR option.",
    "busy": "Our server is working on another file. Please try again in a few seconds.",
    "timeout": "This file took too long to process. Try a smaller file or fewer pages.",
    "server_memory": "This file needed more memory than our server has. Try a smaller file or fewer pages.",
    "rate_limited": "You have sent many files in a short time. Please wait a few minutes and try again.",
    "forbidden_origin": "Requests from this website are not allowed.",
    "bad_request": "Some required information is missing. Please try again.",
    "bad_language": "This OCR language is not available.",
    "generic": "We could not process your file. Please try again.",
}


def file_too_large_message(limit_mb: int) -> str:
    # Same wording as before, so existing pages keep showing a sensible line.
    return f"File too large. Maximum allowed size for this tool is {limit_mb} MB."


class ApiError(Exception):
    """An error with a stable code, a safe message and an HTTP status."""

    def __init__(self, status: int, code: str, message: str | None = None, headers: dict | None = None):
        super().__init__(code)
        self.status = status
        self.code = code
        self.message = message or MESSAGES.get(code, MESSAGES["generic"])
        self.headers = headers or {}

    def payload(self) -> dict:
        return {"code": self.code, "message": self.message, "detail": self.message}


# ---------------------------------------------------------------------------
# File names
# ---------------------------------------------------------------------------

def display_base(filename: str | None, default: str = "document") -> str:
    """Clean the user's file name for use in a download name only.

    Drops any folder part, the extension, and every character that is not a
    letter, digit, space, dot, dash or underscore. Limits it to 100 characters.
    It is never used to build a path on disk.
    """
    name = (filename or "").replace("\\", "/").rsplit("/", 1)[-1]
    base = os.path.splitext(name)[0] if not name.startswith(".") else name.lstrip(".")
    base = re.sub(r"[^\w .-]", "", base, flags=re.UNICODE)
    base = re.sub(r"\s+", " ", base).strip(" .-")
    return base[:100].strip(" .-") or default


def safe_extension(filename: str | None, default: str) -> str:
    """Return a short lower case extension such as '.docx', or the default."""
    name = (filename or "").replace("\\", "/").rsplit("/", 1)[-1]
    ext = os.path.splitext(name)[1].lower()
    return ext if re.fullmatch(r"\.[a-z0-9]{1,5}", ext) else default


# ---------------------------------------------------------------------------
# Worker tasks. These run only inside the child process, so the heavy
# libraries are imported here and never load into the API process.
# Every task reads <job_dir>/input<ext> and writes a fixed output name.
# ---------------------------------------------------------------------------

def _check_pdf(path: str, encrypted_status: int, reject_any_encryption: bool = False) -> int:
    """Open a PDF to check it. Returns the page count."""
    import pikepdf

    try:
        with pikepdf.open(path) as pdf:
            if reject_any_encryption and pdf.is_encrypted:
                raise ApiError(encrypted_status, "encrypted")
            return len(pdf.pages)
    except pikepdf.PasswordError:
        raise ApiError(encrypted_status, "encrypted")
    except ApiError:
        raise
    except Exception:
        raise ApiError(422, "corrupt")


def _check_page_cap(pages: int, cap: int) -> None:
    if cap and pages > cap:
        raise ApiError(
            413,
            "too_many_pages",
            f"This file has {pages} pages. The limit for this tool is {cap} pages.",
        )


def _require_output(path: str) -> None:
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        raise ApiError(500, "generic")


def task_pdf_to_word(job_dir: str, params: dict) -> dict:
    src = os.path.join(job_dir, "input.pdf")
    out = os.path.join(job_dir, "output.docx")

    pages = _check_pdf(src, 403)
    _check_page_cap(pages, int(params.get("max_pages") or 0))

    from pdf2docx import Converter

    converter = Converter(src)
    try:
        converter.convert(
            out,
            start=0,
            end=None,
            multi_processing=False,
            page_layout_analysis=True,
        )
    finally:
        converter.close()

    _require_output(out)
    return {}


# Image modes that carry an alpha channel of their own.
TRANSPARENT_MODES = {"RGBA", "LA", "La", "PA", "RGBa"}

# Modes we do not re-encode, to avoid a colour shift. JPEG has no CMYK path
# that survives a round trip through Pillow without changing the colours.
NO_RECODE_MODES = {"CMYK", "YCCK"}


def _image_stream(img):
    """The PDF image XObject behind a pypdf image, or None if it is not reachable."""
    try:
        reference = getattr(img, "indirect_reference", None)
        return reference.get_object() if reference is not None else None
    except Exception:
        return None


def _is_masked(stream) -> bool:
    """True when the PDF itself gives this image a transparency mask.

    A signature or logo placed by Sign PDF or Watermark PDF arrives as an image
    plus an /SMask holding its alpha channel. Re-encoding the image as JPEG
    throws the mask away, which is what turned a signature into a solid box.
    """
    if stream is None:
        return False
    for key in ("/SMask", "/Mask"):
        try:
            if key in stream and stream.get(key) is not None:
                return True
        except Exception:
            # An unreadable dictionary is treated as masked, so the image is
            # left alone rather than risked.
            return True
    return False


def _keep_as_is(img, stream) -> bool:
    """True when an image must be copied across untouched.

    Checked before the image is decoded where possible, so a skipped image
    costs nothing.
    """
    if _is_masked(stream):
        return True
    try:
        pil_img = img.image
    except Exception:
        return True
    if pil_img.mode in TRANSPARENT_MODES or pil_img.mode in NO_RECODE_MODES:
        return True
    # A palette or grayscale image can carry transparency in its own metadata.
    if "transparency" in getattr(pil_img, "info", {}):
        return True
    return False


def task_compress(job_dir: str, params: dict) -> dict:
    from PIL import Image
    from pypdf import PdfReader, PdfWriter

    src = os.path.join(job_dir, "input.pdf")
    out = os.path.join(job_dir, "output.pdf")

    try:
        reader = PdfReader(src)
        encrypted = reader.is_encrypted
    except Exception:
        raise ApiError(422, "corrupt")
    if encrypted:
        raise ApiError(403, "encrypted")

    level = params.get("level", "medium")
    if level == "extreme":
        img_quality, scale_factor = 20, 0.5
    elif level == "medium":
        img_quality, scale_factor = 40, 0.75
    else:
        img_quality, scale_factor = 70, 1.0

    writer = PdfWriter(clone_from=src)
    for page in writer.pages:
        for img in page.images:
            try:
                # An image with transparency, or one in CMYK, is left exactly
                # as it is. JPEG cannot hold an alpha channel, so re-encoding a
                # signature or a logo would flatten it to a solid box, and a
                # CMYK round trip shifts the colours.
                if _keep_as_is(img, _image_stream(img)):
                    continue
                pil_img = img.image
                if pil_img.mode == "P":
                    pil_img = pil_img.convert("RGB")
                if scale_factor < 1.0:
                    width, height = pil_img.size
                    new_size = (max(1, int(width * scale_factor)), max(1, int(height * scale_factor)))
                    pil_img = pil_img.resize(new_size, Image.Resampling.LANCZOS)
                img.replace(pil_img, quality=img_quality)
            except Exception:
                # Same as before: an image that cannot be re-encoded is left as it is.
                continue

    with open(out, "wb") as handle:
        writer.write(handle)

    # Same as before: never return a file larger than the original.
    if os.path.getsize(out) >= os.path.getsize(src):
        shutil.copyfile(src, out)

    _require_output(out)
    return {}


def task_protect(job_dir: str, params: dict) -> dict:
    import pikepdf

    src = os.path.join(job_dir, "input.pdf")
    out = os.path.join(job_dir, "output.pdf")

    clean_password = (params.get("password") or "").strip()
    if not clean_password:
        raise ApiError(400, "bad_request", "Password cannot be empty.")

    try:
        with pikepdf.open(src) as pdf:
            if pdf.is_encrypted:
                raise ApiError(400, "encrypted", MESSAGES["already_encrypted"])
            pdf.save(
                out,
                encryption=pikepdf.Encryption(user=clean_password, owner=clean_password, R=6),
            )
    except pikepdf.PasswordError:
        raise ApiError(400, "encrypted", MESSAGES["already_encrypted"])
    except ApiError:
        raise
    except Exception:
        raise ApiError(422, "corrupt")

    _require_output(out)
    return {}


def task_unlock(job_dir: str, params: dict) -> dict:
    import pikepdf
    from pypdf import PdfReader, PdfWriter

    src = os.path.join(job_dir, "input.pdf")
    out = os.path.join(job_dir, "output.pdf")

    password = params.get("password") or ""
    clean = password.strip()

    # Same order as before: trimmed password, raw password, then empty.
    passwords_to_try = []
    if clean:
        passwords_to_try.append(clean)
    if password and password != clean:
        passwords_to_try.append(password)
    if "" not in passwords_to_try:
        passwords_to_try.append("")

    saw_password_error = False

    # Tier 1: pikepdf. Saving without an encryption argument removes both the
    # open password and the print, copy and edit limits, and keeps bookmarks.
    for pwd in passwords_to_try:
        try:
            with pikepdf.open(src, password=pwd) as pdf:
                pdf.save(out)
            _require_output(out)
            return {}
        except pikepdf.PasswordError:
            saw_password_error = True
            continue
        except Exception:
            continue

    # Tier 2: pypdf fallback, as before.
    try:
        reader = PdfReader(src)
        for pwd in passwords_to_try:
            try:
                result = reader.decrypt(pwd) if reader.is_encrypted else 1
                if result:
                    writer = PdfWriter()
                    for page in reader.pages:
                        writer.add_page(page)
                    with open(out, "wb") as handle:
                        writer.write(handle)
                    _require_output(out)
                    return {}
            except Exception:
                continue
        if reader.is_encrypted:
            saw_password_error = True
    except Exception:
        pass

    if not saw_password_error:
        raise ApiError(422, "corrupt")
    if clean or password:
        raise ApiError(401, "wrong_password")
    raise ApiError(401, "wrong_password", MESSAGES["password_needed"])


def _ocrmypdf_error_to_api(exc: Exception, encrypted_status: int) -> ApiError:
    from ocrmypdf import exceptions as ocr_exc

    def is_a(name: str) -> bool:
        cls = getattr(ocr_exc, name, None)
        return cls is not None and isinstance(exc, cls)

    if is_a("EncryptedPdfError"):
        return ApiError(encrypted_status, "encrypted")
    if is_a("InputFileError"):
        return ApiError(422, "corrupt")
    if is_a("MissingDependencyError") or is_a("BadArgsError"):
        return ApiError(400, "bad_language")
    return ApiError(500, "generic")


def task_ocr(job_dir: str, params: dict) -> dict:
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
    return {}


def task_pdf_to_excel(job_dir: str, params: dict) -> dict:
    import pandas as pd
    import pdfplumber

    src = os.path.join(job_dir, "input.pdf")
    out = os.path.join(job_dir, "output.xlsx")

    pages = _check_pdf(src, 400)

    target = src
    if params.get("ocr"):
        _check_page_cap(pages, int(params.get("max_pages_ocr") or 0))
        ocr_out = os.path.join(job_dir, "ocr.pdf")
        try:
            import ocrmypdf

            ocrmypdf.ocr(
                src,
                ocr_out,
                language="eng",
                force_ocr=True,
                optimize=0,
                jobs=1,
                skip_big=10.0,
            )
            if os.path.exists(ocr_out) and os.path.getsize(ocr_out) > 0:
                target = ocr_out
        except MemoryError:
            raise
        except Exception:
            # Same as before: if OCR fails, extract from the original file.
            target = src

    all_rows = []
    with pdfplumber.open(target) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                tables = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_y_tolerance": 15,
                })
            for table in tables:
                for row in table:
                    cleaned = [cell.replace("\n", " ") if cell is not None else "" for cell in row]
                    if any(cell.strip() for cell in cleaned):
                        all_rows.append(cleaned)

    if not all_rows:
        raise ApiError(400, "no_tables")

    frame = pd.DataFrame(all_rows)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Table_Data", index=False, header=False)

    _require_output(out)
    return {}


def task_excel_prepare(job_dir: str, params: dict) -> dict:
    """Same formatting as before: landscape and one page wide for every sheet.

    Returns the name of the file LibreOffice should convert. If anything fails,
    the original upload is converted unchanged, as before.
    """
    ext = params.get("ext", ".xlsx")
    original = f"input{ext}"
    src = os.path.join(job_dir, original)
    target_name = original

    try:
        import pandas as pd
        from openpyxl import load_workbook
        from openpyxl.worksheet.properties import PageSetupProperties

        if ext == ".csv":
            frame = pd.read_csv(src)
            target_name = "prepared.xlsx"
            frame.to_excel(os.path.join(job_dir, target_name), index=False)

        if target_name.endswith(".xlsx"):
            target = os.path.join(job_dir, target_name)
            workbook = load_workbook(target)
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                sheet.page_setup.orientation = "landscape"
                sheet.page_setup.fitToPage = True
                sheet.page_setup.fitToHeight = 0
                sheet.page_setup.fitToWidth = 1
                if sheet.sheet_properties.pageSetUpPr is None:
                    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
                else:
                    sheet.sheet_properties.pageSetUpPr.fitToPage = True
            workbook.save(target)
    except MemoryError:
        raise
    except Exception:
        target_name = original

    return {"target": target_name}


WORKER_TASKS = {
    "pdf_to_word": task_pdf_to_word,
    "compress": task_compress,
    "protect": task_protect,
    "unlock": task_unlock,
    "ocr": task_ocr,
    "pdf_to_excel": task_pdf_to_excel,
    "excel_prepare": task_excel_prepare,
}


def _worker_main(task: str, job_dir: str) -> int:
    """Entry point of the child process. Writes <job_dir>/result.json."""
    try:
        params = json.loads(sys.stdin.read() or "{}")
    except ValueError:
        params = {}

    try:
        extra = WORKER_TASKS[task](job_dir, params) or {}
        result = {"ok": True, **extra}
    except ApiError as err:
        result = {"ok": False, "status": err.status, "code": err.code, "message": err.message}
    except MemoryError:
        result = {"ok": False, "status": 422, "code": "server_memory"}
    except Exception as exc:
        # Only the exception type is passed on, never its text.
        result = {"ok": False, "status": 500, "code": "generic", "error_type": type(exc).__name__}

    tmp_path = os.path.join(job_dir, "result.json.tmp")
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(result, handle)
    os.replace(tmp_path, os.path.join(job_dir, "result.json"))
    return 0


if __name__ == "__main__" and len(sys.argv) >= 4 and sys.argv[1] == "--worker":
    sys.exit(_worker_main(sys.argv[2], sys.argv[3]))


# ---------------------------------------------------------------------------
# Everything below runs in the API process only.
# ---------------------------------------------------------------------------

from contextlib import asynccontextmanager  # noqa: E402

from fastapi import FastAPI, File, Form, Request, UploadFile  # noqa: E402
from fastapi.exceptions import RequestValidationError  # noqa: E402
from fastapi.middleware.cors import CORSMiddleware  # noqa: E402
from fastapi.responses import FileResponse, JSONResponse  # noqa: E402
from starlette.background import BackgroundTask  # noqa: E402

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("optimax")


# ---------------------------------------------------------------------------
# Child processes
# ---------------------------------------------------------------------------

# Raise the child's out of memory score so the kernel stops the job, not the API.
_OOM_PREFIX = ["sh", "-c", '{ echo 1000 > /proc/self/oom_score_adj; } 2>/dev/null; exec "$@"', "sh"]


def _child_env() -> dict:
    env = dict(os.environ)
    env.setdefault("OMP_THREAD_LIMIT", "1")
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    return env


def _kill_group(proc: subprocess.Popen) -> None:
    try:
        os.killpg(proc.pid, signal.SIGKILL)
    except (ProcessLookupError, PermissionError, OSError):
        pass
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        pass


def run_child(cmd: list, timeout: float, cwd: str, stdin_bytes: bytes | None = None) -> int:
    """Run a command in its own process group with a hard timeout.

    Output is discarded so file contents and names never reach the logs.
    Raises ApiError(timeout) and kills the whole group when time runs out.
    """
    proc = subprocess.Popen(
        _OOM_PREFIX + list(cmd),
        cwd=cwd,
        stdin=subprocess.PIPE if stdin_bytes is not None else subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        start_new_session=True,
        env=_child_env(),
    )
    try:
        proc.communicate(input=stdin_bytes, timeout=max(1.0, timeout))
    except subprocess.TimeoutExpired:
        _kill_group(proc)
        raise ApiError(422, "timeout")
    finally:
        # Stop anything the command left behind in its group.
        try:
            os.killpg(proc.pid, signal.SIGKILL)
        except (ProcessLookupError, PermissionError, OSError):
            pass
    return proc.returncode


# ---------------------------------------------------------------------------
# Job slots: one job at a time on the free tier, with a short queue
# ---------------------------------------------------------------------------

class SlotPool:
    def __init__(self, size: int, max_waiting: int):
        self._free: queue.Queue = queue.Queue()
        for slot in range(size):
            self._free.put(slot)
        self._lock = threading.Lock()
        self._waiting = 0
        self._max_waiting = max_waiting

    def acquire(self, timeout: float) -> int:
        try:
            return self._free.get_nowait()
        except queue.Empty:
            pass
        with self._lock:
            if self._waiting >= self._max_waiting:
                raise ApiError(503, "busy", headers={"Retry-After": str(BUSY_RETRY_AFTER_SECONDS)})
            self._waiting += 1
        try:
            return self._free.get(timeout=timeout)
        except queue.Empty:
            raise ApiError(503, "busy", headers={"Retry-After": str(BUSY_RETRY_AFTER_SECONDS)})
        finally:
            with self._lock:
                self._waiting -= 1

    def release(self, slot: int) -> None:
        self._free.put(slot)


SLOTS = SlotPool(MAX_JOBS, MAX_WAITING)


# ---------------------------------------------------------------------------
# Temp folders
# ---------------------------------------------------------------------------

def remove_dir(path: str | None) -> None:
    if path:
        shutil.rmtree(path, ignore_errors=True)


_sweep_lock = threading.Lock()
_last_sweep = 0.0


def sweep_stale_jobs(force: bool = False) -> None:
    """Delete job folders older than STALE_JOB_SECONDS.

    Covers the rare case where a download was cut off and the normal cleanup
    after sending never ran. Runs at most once a minute.
    """
    global _last_sweep
    now = time.time()
    with _sweep_lock:
        if not force and now - _last_sweep < 60:
            return
        _last_sweep = now
    try:
        entries = list(os.scandir(JOBS_ROOT))
    except FileNotFoundError:
        return
    for entry in entries:
        try:
            if now - entry.stat().st_mtime > STALE_JOB_SECONDS:
                remove_dir(entry.path)
        except OSError:
            continue


def startup_cleanup() -> None:
    """Delete everything left by a previous run, then recreate the folders."""
    for root in (JOBS_ROOT, LO_PROFILES_ROOT):
        remove_dir(root)
        os.makedirs(root, exist_ok=True)


def upload_size(upload: UploadFile) -> int:
    handle = upload.file
    handle.seek(0, os.SEEK_END)
    size = handle.tell()
    handle.seek(0)
    return size


# ---------------------------------------------------------------------------
# A job: one request from upload to response
# ---------------------------------------------------------------------------

class Job:
    def __init__(self, path: str, upload: UploadFile):
        self.tool = path.rsplit("/", 1)[-1]
        self.max_mb, base_timeout = TOOL_LIMITS[path]
        self.timeout = base_timeout * TIMEOUT_FACTOR
        self.upload = upload
        self.id = uuid.uuid4().hex[:12]
        self.base = display_base(upload.filename)
        self.dir: str | None = None
        self.slot: int | None = None
        self.deadline = 0.0
        self.responded = False
        self.error_type: str | None = None
        self.started = time.monotonic()

    def __enter__(self) -> "Job":
        sweep_stale_jobs()
        size = upload_size(self.upload)
        if size > self.max_mb * MB:
            raise ApiError(413, "file_too_large", file_too_large_message(self.max_mb))
        if size == 0:
            raise ApiError(400, "empty_file")
        os.makedirs(JOBS_ROOT, exist_ok=True)
        self.dir = tempfile.mkdtemp(prefix="job-", dir=JOBS_ROOT)
        return self

    def start(self, ext: str) -> None:
        """Save the upload under a fixed name, then wait for a free slot."""
        target = os.path.join(self.dir, f"input{ext}")
        self.upload.file.seek(0)
        with open(target, "wb") as handle:
            shutil.copyfileobj(self.upload.file, handle, length=1 * MB)
        self.slot = SLOTS.acquire(QUEUE_WAIT_SECONDS)
        self.deadline = time.monotonic() + self.timeout

    def _remaining(self) -> float:
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise ApiError(422, "timeout")
        return remaining

    def run_worker(self, task: str, params: dict) -> dict:
        result_path = os.path.join(self.dir, "result.json")
        if os.path.exists(result_path):
            os.remove(result_path)

        returncode = run_child(
            [sys.executable, WORKER_FILE, "--worker", task, self.dir],
            self._remaining(),
            cwd=self.dir,
            stdin_bytes=json.dumps(params).encode("utf-8"),
        )

        try:
            with open(result_path, encoding="utf-8") as handle:
                result = json.load(handle)
        except (OSError, ValueError):
            result = None

        if result is None:
            # No result means the child died. SIGKILL here is almost always the
            # kernel stopping it for using too much memory.
            if returncode in (-signal.SIGKILL, 137):
                raise ApiError(422, "server_memory")
            self.error_type = f"exit:{returncode}"
            raise ApiError(500, "generic")

        if not result.get("ok"):
            self.error_type = result.get("error_type")
            raise ApiError(
                int(result.get("status", 500)),
                result.get("code", "generic"),
                result.get("message"),
            )
        return result

    def libreoffice_to_pdf(self, src_name: str) -> None:
        """Convert <job_dir>/<src_name> to <job_dir>/output.pdf."""
        out_dir = os.path.join(self.dir, "lo-out")
        os.makedirs(out_dir, exist_ok=True)
        profile = os.path.join(LO_PROFILES_ROOT, f"slot-{self.slot}")
        cmd = [
            LIBREOFFICE_BIN,
            f"-env:UserInstallation=file://{profile}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--convert-to",
            "pdf",
            "--outdir",
            out_dir,
            os.path.join(self.dir, src_name),
        ]
        try:
            run_child(cmd, self._remaining(), cwd=self.dir)
        except ApiError:
            # A killed LibreOffice can leave a lock in its profile. Start fresh next time.
            remove_dir(profile)
            raise

        produced = [name for name in os.listdir(out_dir) if name.lower().endswith(".pdf")]
        produced_path = os.path.join(out_dir, produced[0]) if produced else None
        if not produced_path or os.path.getsize(produced_path) == 0:
            remove_dir(profile)
            raise ApiError(422, "convert_failed")
        os.replace(produced_path, os.path.join(self.dir, "output.pdf"))

    def file_response(self, name: str, media_type: str, download_name: str) -> FileResponse:
        path = os.path.join(self.dir, name)
        if not os.path.exists(path) or os.path.getsize(path) == 0:
            raise ApiError(500, "generic")
        self.responded = True
        return FileResponse(
            path,
            media_type=media_type,
            filename=download_name,
            background=BackgroundTask(remove_dir, self.dir),
        )

    def __exit__(self, exc_type, exc, tb):
        if self.slot is not None:
            SLOTS.release(self.slot)
            self.slot = None

        if exc_type is not None or not self.responded:
            remove_dir(self.dir)

        elapsed_ms = int((time.monotonic() - self.started) * 1000)
        if exc_type is None:
            outcome = "ok"
        elif isinstance(exc, ApiError):
            outcome = exc.code
        else:
            outcome = "generic"
            self.error_type = exc_type.__name__
        extra = f" error_type={self.error_type}" if self.error_type else ""
        # No file names, paths or passwords in the log.
        log.info("job=%s tool=%s result=%s ms=%d%s", self.id, self.tool, outcome, elapsed_ms, extra)

        if exc_type is not None and not isinstance(exc, ApiError):
            # Turn any unexpected error into a safe reply that still gets CORS headers.
            raise ApiError(500, "generic") from None
        return False


# ---------------------------------------------------------------------------
# Request guard: allowed origins, early size check, rate limit
# ---------------------------------------------------------------------------

class RateLimiter:
    def __init__(self, limit: int, window_seconds: int):
        self.limit = limit
        self.window = window_seconds
        self._hits: dict[str, deque] = {}
        self._lock = threading.Lock()

    def check(self, key: str) -> tuple[bool, int]:
        if self.limit <= 0:
            return True, 0
        now = time.monotonic()
        with self._lock:
            hits = self._hits.setdefault(key, deque())
            while hits and now - hits[0] > self.window:
                hits.popleft()
            if len(hits) >= self.limit:
                return False, int(self.window - (now - hits[0])) + 1
            hits.append(now)
            if len(self._hits) > 10000:
                for stale_key in [k for k, v in self._hits.items() if not v or now - v[-1] > self.window]:
                    del self._hits[stale_key]
            return True, 0


RATE_LIMITER = RateLimiter(RATE_LIMIT_JOBS, RATE_LIMIT_WINDOW_SECONDS)


def client_ip(scope: dict, headers: dict) -> str:
    # Render sits behind Cloudflare, which sets cf-connecting-ip itself.
    for name in ("cf-connecting-ip",):
        value = headers.get(name)
        if value:
            return value.strip()
    forwarded = headers.get("x-forwarded-for")
    if forwarded:
        return forwarded.split(",")[0].strip()
    client = scope.get("client")
    return client[0] if client else "unknown"


async def send_json(send, status: int, payload: dict, headers: dict | None = None) -> None:
    body = json.dumps(payload).encode("utf-8")
    raw_headers = [
        (b"content-type", b"application/json"),
        (b"content-length", str(len(body)).encode("latin-1")),
    ]
    for key, value in (headers or {}).items():
        raw_headers.append((key.lower().encode("latin-1"), str(value).encode("latin-1")))
    await send({"type": "http.response.start", "status": status, "headers": raw_headers})
    await send({"type": "http.response.body", "body": body})


async def drain_body(receive, limit: int) -> None:
    total = 0
    while True:
        message = await receive()
        if message["type"] != "http.request":
            return
        total += len(message.get("body", b""))
        if not message.get("more_body") or total > limit:
            return


class _BodyTooLarge(Exception):
    pass


class RequestGuardMiddleware:
    """Checks every POST to /api/ before the upload is stored."""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http" or scope.get("method") != "POST" or not scope.get("path", "").startswith("/api/"):
            await self.app(scope, receive, send)
            return

        headers = {k.decode("latin-1").lower(): v.decode("latin-1") for k, v in scope.get("headers", [])}

        # 1. Other websites may not use the API from their visitors' browsers.
        origin = headers.get("origin")
        if origin and origin.rstrip("/") not in ALLOWED_ORIGINS:
            await send_json(send, 403, ApiError(403, "forbidden_origin").payload())
            return

        # 2. Refuse an oversized upload from its declared size, before reading it.
        limit_mb = TOOL_LIMITS.get(scope["path"], (1, 0))[0]
        max_body = limit_mb * MB + MULTIPART_SLACK_BYTES
        too_large = ApiError(413, "file_too_large", file_too_large_message(limit_mb))
        declared = headers.get("content-length")
        if declared:
            try:
                declared_size = int(declared)
            except ValueError:
                declared_size = 0
            if declared_size > max_body:
                await send_json(send, 413, too_large.payload())
                return

        # 3. Simple per visitor limit.
        allowed, retry_after = RATE_LIMITER.check(client_ip(scope, headers))
        if not allowed:
            await drain_body(receive, max_body)
            await send_json(
                send,
                429,
                ApiError(429, "rate_limited").payload(),
                {"Retry-After": str(retry_after)},
            )
            return

        # 4. Count the bytes as they arrive, in case the size was not declared.
        state = {"bytes": 0, "too_large": False, "replied": False}

        async def limited_receive():
            message = await receive()
            if message["type"] == "http.request":
                state["bytes"] += len(message.get("body", b""))
                if state["bytes"] > max_body:
                    state["too_large"] = True
                    raise _BodyTooLarge()
            return message

        async def guarded_send(message):
            if state["too_large"]:
                if not state["replied"]:
                    state["replied"] = True
                    await send_json(send, 413, too_large.payload())
                return
            await send(message)

        try:
            await self.app(scope, limited_receive, guarded_send)
        except Exception:
            if state["too_large"]:
                if not state["replied"]:
                    state["replied"] = True
                    await send_json(send, 413, too_large.payload())
                return
            raise


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

@asynccontextmanager
async def lifespan(_app):
    startup_cleanup()
    log.info(
        "start max_jobs=%d max_waiting=%d timeout_factor=%.2f origins=%s",
        MAX_JOBS, MAX_WAITING, TIMEOUT_FACTOR, ",".join(ALLOWED_ORIGINS),
    )
    yield


app = FastAPI(title="Optimax PDF API", lifespan=lifespan)

# Added first, so it runs inside CORS and its replies still get CORS headers.
app.add_middleware(RequestGuardMiddleware)
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition", "Retry-After"],
    max_age=600,
)


@app.exception_handler(ApiError)
async def api_error_handler(_request: Request, exc: ApiError):
    return JSONResponse(exc.payload(), status_code=exc.status, headers=exc.headers)


@app.exception_handler(RequestValidationError)
async def validation_error_handler(_request: Request, _exc: RequestValidationError):
    return JSONResponse(ApiError(422, "bad_request").payload(), status_code=422)


PDF_TYPE = "application/pdf"
DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@app.post("/api/convert")
def convert_pdf(file: UploadFile = File(...)):
    with Job("/api/convert", file) as job:
        job.start(".pdf")
        job.run_worker("pdf_to_word", {"max_pages": MAX_PAGES_PDF_TO_WORD})
        return job.file_response("output.docx", DOCX_TYPE, f"OptimaxPDF_Converted_{job.base}.docx")


@app.post("/api/word-to-pdf")
def word_to_pdf(file: UploadFile = File(...)):
    with Job("/api/word-to-pdf", file) as job:
        ext = safe_extension(file.filename, ".docx")
        job.start(ext)
        job.libreoffice_to_pdf(f"input{ext}")
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_WordToPDF_{job.base}.pdf")


@app.post("/api/compress")
def compress_pdf(file: UploadFile = File(...), level: str = Form("medium")):
    with Job("/api/compress", file) as job:
        job.start(".pdf")
        job.run_worker("compress", {"level": level})
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_Compressed_{job.base}.pdf")


@app.post("/api/protect")
def protect_pdf(file: UploadFile = File(...), password: str = Form(...)):
    with Job("/api/protect", file) as job:
        if not password.strip():
            raise ApiError(400, "bad_request", "Password cannot be empty.")
        job.start(".pdf")
        job.run_worker("protect", {"password": password})
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_Protected_{job.base}.pdf")


@app.post("/api/unlock")
def unlock_pdf(file: UploadFile = File(...), password: str = Form("")):
    with Job("/api/unlock", file) as job:
        job.start(".pdf")
        job.run_worker("unlock", {"password": password or ""})
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_Unlocked_{job.base}.pdf")


@app.post("/api/ocr")
def ocr_pdf(file: UploadFile = File(...), language: str = Form("eng")):
    with Job("/api/ocr", file) as job:
        if not OCR_LANGUAGE_PATTERN.fullmatch(language or ""):
            raise ApiError(400, "bad_language")
        job.start(".pdf")
        job.run_worker("ocr", {"language": language, "max_pages": MAX_PAGES_OCR})
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_Searchable_{job.base}.pdf")


@app.post("/api/excel-to-pdf")
def excel_to_pdf(file: UploadFile = File(...)):
    with Job("/api/excel-to-pdf", file) as job:
        ext = safe_extension(file.filename, ".xlsx")
        job.start(ext)
        prepared = job.run_worker("excel_prepare", {"ext": ext})
        target = prepared.get("target") or f"input{ext}"
        if not re.fullmatch(r"(input\.[a-z0-9]{1,5}|prepared\.xlsx)", target):
            target = f"input{ext}"
        job.libreoffice_to_pdf(target)
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_ExcelToPDF_{job.base}.pdf")


@app.post("/api/ppt-to-pdf")
def ppt_to_pdf(file: UploadFile = File(...)):
    with Job("/api/ppt-to-pdf", file) as job:
        ext = safe_extension(file.filename, ".pptx")
        job.start(ext)
        job.libreoffice_to_pdf(f"input{ext}")
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_PPTtoPDF_{job.base}.pdf")


@app.post("/api/pdf-to-excel")
def pdf_to_excel(file: UploadFile = File(...), ocr: bool = Form(False)):
    with Job("/api/pdf-to-excel", file) as job:
        job.start(".pdf")
        job.run_worker("pdf_to_excel", {"ocr": bool(ocr), "max_pages_ocr": MAX_PAGES_PDF_TO_EXCEL_OCR})
        return job.file_response("output.xlsx", XLSX_TYPE, f"OptimaxPDF_Extracted_{job.base}.xlsx")


@app.get("/")
async def health_check():
    # Unchanged reply. The home page pings this to wake the server.
    return {"status": "online", "service": "Optimax PDF Backend is Running"}


@app.get("/health")
async def health():
    # Light check for the wake up state on tool pages. Never waits on a job.
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))
