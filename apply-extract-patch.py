#!/usr/bin/env python3
"""Rebuild PDF to Excel, and add PDF to PowerPoint.

Run it from the folder holding server.py, AFTER apply-office-patch.py:

    python apply-extract-patch.py server.py

It changes nothing unless every anchor it expects is found exactly once, and
running it twice is safe. A copy is written to server.py.bak first.

What it does.

1.  PDF to Excel stops flattening everything into one sheet. It can now put
    one sheet per table or one per page, write CSV instead, take a page
    range, treat the first row as a header, and run OCR in any of the 120
    languages on this server rather than English alone. The old request shape
    still produces the old output, down to the Table_Data sheet name, so
    anything scripted against it keeps working.

2.  A preview mode on the same endpoint. `preview=1` returns JSON describing
    the tables found instead of a file, so the page can show what was
    detected before anybody downloads a spreadsheet that may be wrong. The
    job directory is cleaned up on the way out because nothing is streamed.

3.  A new endpoint, POST /api/pdf-to-ppt. LibreOffice imports the PDF through
    its Draw filter and writes a .pptx whose text is genuinely editable: a
    three page PDF produced three slides each holding real text boxes and
    vector shapes, not pictures. The page's own shape is kept, so a 16:9 PDF
    comes back 16:9 rather than being forced to 4:3.

4.  A page range for that endpoint, done by trimming the PDF with pypdf
    first, because the import filter has no range option of its own.

5.  header_row does something visible. Writing the first row as pandas column
    names and writing it as ordinary data produce a byte identical sheet, so
    on its own that option would be a control that does nothing. It makes the
    row bold, freezes it, and puts a filter button on each column.
"""

import os
import shutil
import sys

MARK = "OPTIMAX_EXTRACT_V1"


# --------------------------------------------------------------------------
# 1. Limits for the new endpoint
# --------------------------------------------------------------------------

LIMITS_OLD = '    "/api/pdf-to-excel": (20, 180),'
LIMITS_NEW = '    "/api/pdf-to-excel": (20, 180),\n    "/api/pdf-to-ppt": (20, 180),'


# --------------------------------------------------------------------------
# 2. A LibreOffice helper for PDF to PPTX
# --------------------------------------------------------------------------

LO_ANCHOR = '''        target = os.path.join(self.dir, "converted.xlsx")
        os.replace(produced_path, target)
        return "converted.xlsx"
'''

LO_NEW = '''        target = os.path.join(self.dir, "converted.xlsx")
        os.replace(produced_path, target)
        return "converted.xlsx"

    def libreoffice_pdf_to_pptx(self, src_name: str) -> None:
        """Convert <job_dir>/<src_name> to <job_dir>/output.pptx.

        The PDF goes in through Draw's PDF import filter and out through
        Impress, which gives slides holding real text boxes and vector
        shapes rather than a picture of each page. The page's own size is
        carried across, so a widescreen PDF stays widescreen.
        """
        out_dir = os.path.join(self.dir, "lo-pptx")
        os.makedirs(out_dir, exist_ok=True)
        profile = os.path.join(LO_PROFILES_ROOT, f"slot-{self.slot}")
        cmd = [
            LIBREOFFICE_BIN,
            f"-env:UserInstallation=file://{profile}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--infilter=impress_pdf_import",
            "--convert-to",
            "pptx:Impress Office Open XML",
            "--outdir",
            out_dir,
            os.path.join(self.dir, src_name),
        ]
        try:
            run_child(cmd, self._remaining(), cwd=self.dir)
        except ApiError:
            remove_dir(profile)
            raise

        produced = [name for name in os.listdir(out_dir) if name.lower().endswith(".pptx")]
        produced_path = os.path.join(out_dir, produced[0]) if produced else None
        if not produced_path or os.path.getsize(produced_path) == 0:
            remove_dir(profile)
            raise ApiError(422, "convert_failed")
        os.replace(produced_path, os.path.join(self.dir, "output.pptx"))
'''


# --------------------------------------------------------------------------
# 3. The new workers
# --------------------------------------------------------------------------

WORKERS_NEW = '''MAX_PREVIEW_TABLES = 20
MAX_PREVIEW_ROWS = 5
MAX_PREVIEW_COLUMNS = 8
EXTRACT_SHEET_MODES = ("one", "page", "table")
EXTRACT_OUTPUTS = ("xlsx", "csv")
EXTRACT_MAX_LANGUAGES = 3
EXTRACT_LANG_RE = re.compile(r"^[a-z]{3}(_[a-z]{4,5})?$")
EXTRACT_RANGE_RE = re.compile(r"^[0-9]+(-[0-9]+)?(,[0-9]+(-[0-9]+)?)*$")


def expand_page_range(text, total):
    """"1-3,7" into [1, 2, 3, 7], bounded by the document.

    None means every page, which is also what an unreadable range gives: a
    malformed box converts the whole file rather than failing the job.
    """
    squeezed = re.sub(r"\\s+", "", str(text or ""))
    if not squeezed or len(squeezed) > 200 or not EXTRACT_RANGE_RE.match(squeezed):
        return None
    wanted = set()
    for part in squeezed.split(","):
        if "-" in part:
            first, last = part.split("-", 1)
            first, last = int(first), int(last)
            if first > last:
                first, last = last, first
        else:
            first = last = int(part)
        for number in range(max(1, first), min(int(total), last) + 1):
            wanted.add(number)
    return sorted(wanted) or None


def clean_ocr_languages(text):
    """A Tesseract language string we are willing to run, or English."""
    parts = [p for p in re.split(r"[+,\\s]+", str(text or "").lower()) if p]
    good = [p for p in parts if EXTRACT_LANG_RE.match(p)][:EXTRACT_MAX_LANGUAGES]
    return "+".join(good) if good else "eng"


def safe_sheet_name(name, used):
    """Excel refuses [ ] : * ? / \\\\ and anything over 31 characters."""
    clean = re.sub(r"[\\[\\]:*?/\\\\]", " ", str(name)).strip()[:31] or "Sheet"
    candidate = clean
    counter = 2
    while candidate.lower() in used:
        suffix = f" {counter}"
        candidate = clean[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(candidate.lower())
    return candidate


def _unique_header(row):
    """pandas refuses duplicate column names; a PDF header often has them."""
    out, seen = [], {}
    for index, cell in enumerate(row):
        name = str(cell).strip() or f"Column {index + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name} {seen[name]}"
        else:
            seen[name] = 1
        out.append(name)
    return out


def _collect_tables(path, page_numbers):
    """Every table found, as (page number, rows). The two strategies, and the
    cleaning, are exactly what this tool has always used."""
    import pdfplumber

    found = []
    with pdfplumber.open(path) as pdf:
        for index, page in enumerate(pdf.pages, 1):
            if page_numbers and index not in page_numbers:
                continue
            tables = page.extract_tables()
            if not tables:
                tables = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_y_tolerance": 15,
                })
            for table in tables:
                rows = []
                for row in table:
                    cleaned = [cell.replace("\\n", " ") if cell is not None else "" for cell in row]
                    if any(cell.strip() for cell in cleaned):
                        rows.append(cleaned)
                if rows:
                    found.append((index, rows))
    return found


def _group_tables(found, mode):
    """[(sheet name, rows)] for the chosen layout."""
    if mode == "table":
        return [(f"Table {i}", rows) for i, (_, rows) in enumerate(found, 1)]
    if mode == "page":
        by_page = {}
        for page, rows in found:
            if page in by_page:
                by_page[page].append([])  # a blank line between tables
            else:
                by_page[page] = []
            by_page[page].extend(rows)
        return [(f"Page {page}", rows) for page, rows in sorted(by_page.items())]
    # One sheet, which is what this tool has always produced. The sheet name
    # is kept so anything scripted against the old output still finds it.
    flat = []
    for _, rows in found:
        flat.extend(rows)
    return [("Table_Data", flat)]


def _write_tables_xlsx(path, groups, header_row):
    import pandas as pd
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        used = set()
        for name, rows in groups:
            if not rows:
                continue
            sheet_name = safe_sheet_name(name, used)
            as_header = bool(header_row) and len(rows) > 1
            if as_header:
                frame = pd.DataFrame(rows[1:], columns=_unique_header(rows[0]))
                frame.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            sheet = writer.sheets[sheet_name]
            width = max((len(row) for row in rows), default=0)
            if as_header and width:
                # Writing the first row as column names and writing it as
                # data give a byte identical sheet, so without this the
                # option would be a control that does nothing.
                for cell in sheet[1]:
                    cell.font = Font(bold=True)
                sheet.freeze_panes = "A2"
                sheet.auto_filter.ref = f"A1:{get_column_letter(width)}{sheet.max_row}"
            for index in range(1, width + 1):
                longest = max(
                    (len(str(row[index - 1])) for row in rows if index <= len(row)),
                    default=0,
                )
                sheet.column_dimensions[get_column_letter(index)].width = min(60, max(9, longest + 2))


class _CsvSink:
    """csv.writer needs something with write(); this keeps it in memory."""

    def __init__(self, store):
        self.store = store

    def write(self, text):
        self.store.append(text)


def _write_tables_csv(job_dir, groups):
    """One CSV, or a ZIP of them. Returns the file name written."""
    import csv
    import zipfile

    if len(groups) <= 1:
        out = os.path.join(job_dir, "output.csv")
        # utf-8-sig, so Excel on Windows opens accented text correctly.
        with open(out, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            for row in (groups[0][1] if groups else []):
                writer.writerow(row)
        return "output.csv"

    out = os.path.join(job_dir, "output.zip")
    used = set()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as bundle:
        for name, rows in groups:
            if not rows:
                continue
            buffer = []
            writer = csv.writer(_CsvSink(buffer))
            for row in rows:
                writer.writerow(row)
            bundle.writestr(f"{safe_sheet_name(name, used)}.csv", "".join(buffer).encode("utf-8-sig"))
    return "output.zip"


def task_pdf_to_excel(job_dir: str, params: dict) -> dict:
    src = os.path.join(job_dir, "input.pdf")
    pages = _check_pdf(src, 400)

    mode = str(params.get("sheets") or "one").lower()
    if mode not in EXTRACT_SHEET_MODES:
        mode = "one"
    output = str(params.get("output") or "xlsx").lower()
    if output not in EXTRACT_OUTPUTS:
        output = "xlsx"
    header_row = str(params.get("header_row") or "").strip().lower() in ("1", "true", "yes", "on")
    wanted_pages = expand_page_range(params.get("pages"), pages)
    preview_only = str(params.get("preview") or "").strip().lower() in ("1", "true", "yes", "on")

    target = src
    if params.get("ocr"):
        _check_page_cap(pages, int(params.get("max_pages_ocr") or 0))
        ocr_out = os.path.join(job_dir, "ocr.pdf")
        try:
            import ocrmypdf

            ocrmypdf.ocr(
                src,
                ocr_out,
                language=clean_ocr_languages(params.get("languages")),
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

    found = _collect_tables(target, wanted_pages)

    if preview_only:
        # No file is written, so the job does not stream anything and its
        # directory is cleared on the way out.
        return {
            "preview": {
                "pages": pages,
                "truncated": len(found) > MAX_PREVIEW_TABLES,
                "tables": [
                    {
                        "page": page,
                        "total_rows": len(rows),
                        "total_columns": max((len(r) for r in rows), default=0),
                        "rows": [row[:MAX_PREVIEW_COLUMNS] for row in rows[:MAX_PREVIEW_ROWS]],
                    }
                    for page, rows in found[:MAX_PREVIEW_TABLES]
                ],
            }
        }

    if not found:
        raise ApiError(400, "no_tables")

    groups = _group_tables(found, mode)
    if output == "csv":
        name = _write_tables_csv(job_dir, groups)
    else:
        name = "output.xlsx"
        _write_tables_xlsx(os.path.join(job_dir, name), groups, header_row)

    _require_output(os.path.join(job_dir, name))
    return {"file": name, "tables": len(found)}


def task_pdf_subset(job_dir: str, params: dict) -> dict:
    """Write <job_dir>/subset.pdf holding only the pages asked for.

    PDF to PowerPoint needs this because the Draw import filter has no page
    range option of its own. If the range cannot be read, nothing is written
    and the whole document is converted.
    """
    from pypdf import PdfReader, PdfWriter

    src = os.path.join(job_dir, "input.pdf")
    try:
        reader = PdfReader(src)
    except Exception:
        raise ApiError(422, "corrupt")
    if reader.is_encrypted:
        raise ApiError(403, "encrypted")

    wanted = expand_page_range(params.get("pages"), len(reader.pages))
    if not wanted:
        return {}

    writer = PdfWriter()
    for number in wanted:
        writer.add_page(reader.pages[number - 1])
    out = os.path.join(job_dir, "subset.pdf")
    with open(out, "wb") as handle:
        writer.write(handle)
    _require_output(out)
    return {"target": "subset.pdf"}


'''

WORKER_MAP_OLD = '    "pdf_to_excel": task_pdf_to_excel,'
WORKER_MAP_NEW = '    "pdf_to_excel": task_pdf_to_excel,\n    "pdf_subset": task_pdf_subset,'


# --------------------------------------------------------------------------
# 4. The media type for the new output
# --------------------------------------------------------------------------

TYPE_OLD = 'XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"'
TYPE_NEW = (
    'XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"\n'
    'PPTX_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"\n'
    'CSV_TYPE = "text/csv"'
)


# --------------------------------------------------------------------------
# 5. The endpoints
# --------------------------------------------------------------------------

ENDPOINT_OLD = '''@app.post("/api/pdf-to-excel")
def pdf_to_excel(file: UploadFile = File(...), ocr: bool = Form(False)):
    with Job("/api/pdf-to-excel", file) as job:
        job.start(".pdf")
        job.run_worker("pdf_to_excel", {"ocr": bool(ocr), "max_pages_ocr": MAX_PAGES_PDF_TO_EXCEL_OCR})
        return job.file_response("output.xlsx", XLSX_TYPE, f"OptimaxPDF_Extracted_{job.base}.xlsx")'''

ENDPOINT_NEW = '''@app.post("/api/pdf-to-excel")
def pdf_to_excel(
    file: UploadFile = File(...),
    ocr: bool = Form(False),
    sheets: str = Form("one"),
    output: str = Form("xlsx"),
    header_row: str = Form(""),
    pages: str = Form(""),
    languages: str = Form(""),
    preview: str = Form(""),
):
    with Job("/api/pdf-to-excel", file) as job:
        job.start(".pdf")
        result = job.run_worker(
            "pdf_to_excel",
            {
                "ocr": bool(ocr),
                "max_pages_ocr": MAX_PAGES_PDF_TO_EXCEL_OCR,
                "sheets": sheets,
                "output": output,
                "header_row": header_row,
                "pages": pages,
                "languages": languages,
                "preview": preview,
            },
        )

        if result.get("preview") is not None:
            # JSON, not a file. Nothing is streamed, so the job's directory
            # is removed on the way out of the context manager.
            return result["preview"]

        name = result.get("file") or "output.xlsx"
        if name == "output.csv":
            return job.file_response("output.csv", CSV_TYPE, f"OptimaxPDF_Extracted_{job.base}.csv")
        if name == "output.zip":
            return job.file_response("output.zip", "application/zip", f"OptimaxPDF_Extracted_{job.base}.zip")
        return job.file_response("output.xlsx", XLSX_TYPE, f"OptimaxPDF_Extracted_{job.base}.xlsx")


@app.post("/api/pdf-to-ppt")
def pdf_to_ppt(file: UploadFile = File(...), pages: str = Form("")):
    """PDF to PowerPoint, with text that can still be edited.

    LibreOffice reads the PDF through Draw's import filter and writes
    Impress's OOXML, which yields real text boxes and vector shapes rather
    than a picture of each page, and keeps the page's own proportions.

    The browser can also build a presentation from rendered page images with
    nothing uploaded at all; that path lives entirely on the frontend and
    never reaches this endpoint.
    """
    with Job("/api/pdf-to-ppt", file) as job:
        job.start(".pdf")
        source = "input.pdf"
        if str(pages or "").strip():
            subset = job.run_worker("pdf_subset", {"pages": pages})
            if subset.get("target") == "subset.pdf":
                source = "subset.pdf"
        job.libreoffice_pdf_to_pptx(source)
        return job.file_response(
            "output.pptx", PPTX_TYPE, f"OptimaxPDF_PDFtoPPT_{job.base}.pptx"
        )'''


STEPS = [
    ("the per tool limits", LIMITS_OLD, LIMITS_NEW),
    ("the LibreOffice helpers", LO_ANCHOR, LO_NEW),
    ("the media types", TYPE_OLD, TYPE_NEW),
    ("the worker task list", WORKER_MAP_OLD, WORKER_MAP_NEW),
    ("the PDF to Excel endpoint", ENDPOINT_OLD, ENDPOINT_NEW),
]


def replace_worker(source: str) -> str:
    """Swap the whole of task_pdf_to_excel for the new workers.

    The end marker is the definition that immediately follows
    task_pdf_to_excel, not the next one this patch happens to know the name
    of. The first version of this function cut as far as task_excel_prepare,
    and thirteen other definitions sat in between, all of which it silently
    deleted: office_flag, _column_letter, excel_sheet_width_points,
    excel_clean_options and the nine EXCEL_* constants. Python only
    complains about a missing name when the line runs, so the file still
    compiled, the build went green, and two endpoints answered 500 on every
    request. Cutting to the next definition is the whole fix.
    """
    start_marker = "def task_pdf_to_excel(job_dir: str, params: dict) -> dict:"
    end_marker = "\ndef office_flag(value) -> bool:"
    if source.count(start_marker) != 1 or source.count(end_marker) != 1:
        raise LookupError("task_pdf_to_excel followed by office_flag, as the office patch leaves them")
    start = source.index(start_marker)
    end = source.index(end_marker)
    if end < start:
        raise LookupError("task_pdf_to_excel in the position this patch expects")
    return source[:start] + WORKERS_NEW + source[end + 1:]


def main() -> int:
    path = sys.argv[1] if len(sys.argv) > 1 else "server.py"
    if not os.path.exists(path):
        print(f"Nothing was changed. {path} is not here. Run this from the folder holding server.py.")
        return 1

    with open(path, encoding="utf-8") as handle:
        original = handle.read()

    if MARK in original:
        print("Nothing was changed. This patch is already in server.py.")
        return 0

    if "OPTIMAX_OFFICE_OPTIONS_V1" not in original:
        print("Nothing was changed. Apply apply-office-patch.py first; this one builds on it.")
        return 1

    updated = original
    try:
        for label, old, new in STEPS:
            if updated.count(old) != 1:
                raise LookupError(label)
            updated = updated.replace(old, new, 1)
        updated = replace_worker(updated)
    except LookupError as err:
        print("Nothing was changed. Could not find:", err)
        print("The file is not the one this patch was written against. Send me your current server.py.")
        return 1

    updated = "# " + MARK + ": PDF to Excel options and PDF to PowerPoint, 22 November 2026.\n" + updated

    try:
        compile(updated, path, "exec")
    except SyntaxError as err:
        print(f"Nothing was changed. The patched file would not compile: line {err.lineno}, {err.msg}")
        return 1

    shutil.copy(path, path + ".bak")
    with open(path, "w", encoding="utf-8") as handle:
        handle.write(updated)

    print("Patched server.py. A copy of the old one is server.py.bak.")
    print()
    print("  PDF to Excel   one sheet per table or per page, CSV output, page")
    print("                 range, a header row that is bold, frozen and")
    print("                 filtered, OCR in any installed language, and a")
    print("                 preview mode that returns the tables as JSON")
    print("  PDF to PPT     new endpoint POST /api/pdf-to-ppt, editable slides")
    print("                 through LibreOffice, with a page range")
    print()
    print("Commit and push, wait for Render, then check:")
    print("  curl -sS -F \"file=@any.pdf\" -F \"preview=1\" \\")
    print("    https://optimaxpdf-docker-api.onrender.com/api/pdf-to-excel")
    print("That should print JSON, not download a file.")
    print()
    print("Then set EXTRACT_OPTIONS_ENABLED to true in src/assets/js/config.js.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
