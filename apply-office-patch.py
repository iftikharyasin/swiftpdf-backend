#!/usr/bin/env python3
"""Add conversion options to Excel to PDF and PowerPoint to PDF.

Run it from the folder holding server.py:

    python apply-office-patch.py server.py

It changes nothing unless every anchor it expects is found exactly once, and
running it twice is safe: the second run reports that the patch is already in
and exits without touching the file. A copy is written to server.py.bak
before anything changes.

What it does, and why each piece is where it is.

1.  The two endpoints take the new form fields. Every field is optional and
    every default reproduces today's behaviour, so an old page talking to the
    new server gets exactly what it gets now.

2.  Excel page setup is done in openpyxl, before LibreOffice sees the file,
    not through LibreOffice's PDF export filter. The filter has a
    SinglePageSheets option that looks like the right tool and is not:
    production runs LibreOffice 7.4 from Debian bookworm, where support for
    it is doubtful, and its own documentation says it ignores hidden sheets
    and print ranges, which would quietly undo the sheet picker.

3.  Automatic orientation is decided per sheet from the real column widths,
    so a workbook with one wide sheet and two narrow ones comes back with the
    wide one landscape and the other two portrait, in a single PDF. That rule
    lives here and nowhere else; the page just sends "auto".

4.  An .xls or .ods is converted to .xlsx first, because openpyxl reads only
    the modern format. Without that step the layout options would work on
    some uploads and silently do nothing on others. If the first pass fails
    the original is converted directly, which is exactly what happens today.

5.  PowerPoint options are LibreOffice export filter options, tested against
    a real deck: notes pages turned 3 slides into 6 pages, only-notes gave 3
    pages carrying the notes text, hidden slides took it from 3 to 4, a range
    of 1-2 gave 2 pages, and PDF/A-2b produced the pdfaid marker.

6.  Handouts are built afterwards from the converted PDF with pypdf, because
    the export filter has no handout option. The decoration is written onto
    the blank sheet first and the slides are merged on top by pypdf itself;
    appending a second content stream by hand produces a file that readers
    report as damaged.

7.  Both per-tool timeouts go from 120 to 180 seconds, matching OCR. Notes
    pages roughly double the rendering work and the .xls path runs
    LibreOffice twice, so the old ceiling is no longer honest.
"""

import os
import re
import shutil
import sys

MARK = "OPTIMAX_OFFICE_OPTIONS_V1"


# --------------------------------------------------------------------------
# 1. Timeouts
# --------------------------------------------------------------------------

LIMITS_OLD = '    "/api/excel-to-pdf": (20, 120),\n    "/api/ppt-to-pdf": (20, 120),'
LIMITS_NEW = '    "/api/excel-to-pdf": (20, 180),\n    "/api/ppt-to-pdf": (20, 180),'


# --------------------------------------------------------------------------
# 2. A second LibreOffice helper, for turning .xls and .ods into .xlsx
# --------------------------------------------------------------------------

LO_ANCHOR = '        os.replace(produced_path, os.path.join(self.dir, "output.pdf"))\n\n    def file_response('

LO_NEW_METHOD = '''        os.replace(produced_path, os.path.join(self.dir, "output.pdf"))

    def libreoffice_to_xlsx(self, src_name: str) -> str:
        """Convert <job_dir>/<src_name> to <job_dir>/converted.xlsx.

        openpyxl reads only the modern format, so an .xls or an .ods has to
        become one before any page setup can be applied to it. Deliberately a
        sibling of libreoffice_to_pdf rather than a refactor of it: this patch
        runs against a live server, and adding a method cannot break the path
        every other tool already depends on.
        """
        out_dir = os.path.join(self.dir, "lo-xlsx")
        os.makedirs(out_dir, exist_ok=True)
        profile = os.path.join(LO_PROFILES_ROOT, f"slot-{self.slot}")
        cmd = [
            LIBREOFFICE_BIN,
            f"-env:UserInstallation=file://{profile}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--convert-to",
            "xlsx:Calc MS Excel 2007 XML",
            "--outdir",
            out_dir,
            os.path.join(self.dir, src_name),
        ]
        try:
            run_child(cmd, self._remaining(), cwd=self.dir)
        except ApiError:
            remove_dir(profile)
            raise

        produced = [name for name in os.listdir(out_dir) if name.lower().endswith(".xlsx")]
        produced_path = os.path.join(out_dir, produced[0]) if produced else None
        if not produced_path or os.path.getsize(produced_path) == 0:
            remove_dir(profile)
            raise ApiError(422, "convert_failed")
        target = os.path.join(self.dir, "converted.xlsx")
        os.replace(produced_path, target)
        return "converted.xlsx"

    def file_response('''


# --------------------------------------------------------------------------
# 3. The new excel_prepare worker, and the handout worker
# --------------------------------------------------------------------------

PREPARE_NEW = '''def office_flag(value) -> bool:
    """A form field that means yes."""
    return str(value or "").strip().lower() in ("1", "true", "yes", "on")


# Excel's column width is in characters of the default font. The usual
# conversion is pixels = width * 7 + 5, and a PDF point is 0.75 of a pixel at
# 96 dpi. It is an estimate, and it only has to be good enough to tell a
# narrow sheet from a wide one.
EXCEL_PX_PER_CHAR = 7.0
EXCEL_PX_PADDING = 5.0
EXCEL_PT_PER_PX = 0.75
EXCEL_DEFAULT_COL_CHARS = 8.43

# Printable width in points: the paper less normal margins. Used by the
# automatic orientation rule and nothing else.
EXCEL_PORTRAIT_PRINTABLE = {"a4": 494.0, "letter": 511.0, "legal": 511.0, "a3": 741.0}
EXCEL_PAPER_CODES = {"a4": "9", "letter": "1", "legal": "5", "a3": "8"}
EXCEL_MARGIN_SETS = {
    "normal": (0.7, 0.7, 0.75, 0.75),
    "narrow": (0.25, 0.25, 0.75, 0.75),
    "wide": (1.0, 1.0, 1.0, 1.0),
    "none": (0.0, 0.0, 0.0, 0.0),
}
EXCEL_FITS = ("width", "page", "actual", "scale")
EXCEL_ORIENTATIONS = ("auto", "portrait", "landscape")


def _column_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def excel_sheet_width_points(sheet) -> float:
    """Roughly how wide this sheet's used columns print, in points."""
    last = sheet.max_column or 0
    if not last:
        return 0.0
    try:
        fallback = float(sheet.sheet_format.defaultColWidth or EXCEL_DEFAULT_COL_CHARS)
    except (TypeError, ValueError):
        fallback = EXCEL_DEFAULT_COL_CHARS
    if fallback <= 0:
        fallback = EXCEL_DEFAULT_COL_CHARS

    total_px = 0.0
    for index in range(1, last + 1):
        dim = sheet.column_dimensions.get(_column_letter(index))
        if dim is not None and getattr(dim, "hidden", False):
            continue
        chars = fallback
        if dim is not None and dim.width:
            try:
                chars = float(dim.width)
            except (TypeError, ValueError):
                chars = fallback
        total_px += chars * EXCEL_PX_PER_CHAR + EXCEL_PX_PADDING
    return total_px * EXCEL_PT_PER_PX


def excel_clean_options(params: dict) -> dict:
    """Only values this module knows about get through. Anything else falls
    back to the default, so a malformed field converts rather than fails."""
    import json as _json

    orientation = str(params.get("orientation") or "auto").lower()
    if orientation not in EXCEL_ORIENTATIONS:
        orientation = "auto"
    fit = str(params.get("fit") or "width").lower()
    if fit not in EXCEL_FITS:
        fit = "width"
    paper = str(params.get("paper") or "a4").lower()
    if paper not in EXCEL_PAPER_CODES:
        paper = "a4"
    margins = str(params.get("margins") or "normal").lower()
    if margins not in EXCEL_MARGIN_SETS:
        margins = "normal"
    try:
        scale = int(float(params.get("scale") or 100))
    except (TypeError, ValueError):
        scale = 100
    scale = max(10, min(400, scale))

    raw = params.get("sheets")
    keep = None
    if isinstance(raw, str) and raw.strip():
        try:
            parsed = _json.loads(raw)
        except ValueError:
            parsed = None
        if isinstance(parsed, list):
            keep = [str(name) for name in parsed if str(name)]
    elif isinstance(raw, list):
        keep = [str(name) for name in raw if str(name)]

    return {
        "orientation": orientation,
        "fit": fit,
        "paper": paper,
        "margins": margins,
        "scale": scale,
        "gridlines": office_flag(params.get("gridlines")),
        "headings": office_flag(params.get("headings")),
        "repeat_header": office_flag(params.get("repeat_header")),
        "keep": keep,
    }


def task_excel_prepare(job_dir: str, params: dict) -> dict:
    """Set the workbook's page setup before LibreOffice converts it.

    Returns the name of the file LibreOffice should convert. If anything
    fails, the original upload is converted unchanged, exactly as before.
    """
    ext = params.get("ext", ".xlsx")
    source = params.get("source") or f"input{ext}"
    if not re.fullmatch(r"(input\\.[a-z0-9]{1,5}|converted\\.xlsx)", str(source)):
        source = f"input{ext}"
    target_name = source
    options = excel_clean_options(params)

    try:
        from openpyxl import load_workbook
        from openpyxl.worksheet.page import PageMargins
        from openpyxl.worksheet.properties import PageSetupProperties

        if ext == ".csv" and target_name == source and source.endswith(".csv"):
            import pandas as pd

            frame = pd.read_csv(os.path.join(job_dir, source))
            target_name = "prepared.xlsx"
            frame.to_excel(os.path.join(job_dir, target_name), index=False)

        if target_name.endswith(".xlsx") or target_name.endswith(".xlsm"):
            path = os.path.join(job_dir, target_name)
            workbook = load_workbook(path)

            keep = options["keep"]
            if keep:
                wanted = [name for name in workbook.sheetnames if name in keep]
                # Never hand LibreOffice an empty workbook, and never drop
                # everything because one name was misspelt.
                if wanted:
                    for name in list(workbook.sheetnames):
                        if name not in wanted and len(workbook.sheetnames) > 1:
                            del workbook[name]

            left, right, top, bottom = EXCEL_MARGIN_SETS[options["margins"]]
            limit = EXCEL_PORTRAIT_PRINTABLE.get(options["paper"], 494.0)
            for sheet in workbook.worksheets:
                setup = sheet.page_setup
                setup.paperSize = EXCEL_PAPER_CODES[options["paper"]]

                orientation = options["orientation"]
                if orientation == "auto":
                    orientation = "landscape" if excel_sheet_width_points(sheet) > limit else "portrait"
                setup.orientation = orientation

                if sheet.sheet_properties.pageSetUpPr is None:
                    sheet.sheet_properties.pageSetUpPr = PageSetupProperties()
                fit = options["fit"]
                want_fit = fit in ("width", "page")
                # Set pageSetUpPr.fitToPage rather than page_setup.fitToPage.
                # The latter is a property that reaches back through a parent
                # link to this same field, and that link is not always there
                # on a workbook read from disk. Everything in this function is
                # wrapped in a fall back to the untouched file, so an
                # AttributeError here would not fail the job: it would quietly
                # convert as if no option had been chosen, which is the worst
                # of the three outcomes. fitToWidth and fitToHeight are plain
                # fields and need no such care.
                sheet.sheet_properties.pageSetUpPr.fitToPage = want_fit
                if want_fit:
                    setup.fitToWidth = 1
                    setup.fitToHeight = 1 if fit == "page" else 0
                    # Excel ignores a leftover scale when fit to page is on;
                    # LibreOffice has been seen to apply it, so it is cleared.
                    setup.scale = None
                else:
                    setup.scale = options["scale"] if fit == "scale" else 100

                sheet.page_margins = PageMargins(left=left, right=right, top=top, bottom=bottom)
                sheet.print_options.gridLines = options["gridlines"]
                sheet.print_options.headings = options["headings"]
                if options["repeat_header"]:
                    sheet.print_title_rows = "1:1"

            workbook.save(path)
    except MemoryError:
        raise
    except Exception:
        target_name = source

    return {"target": target_name}


# Slides per handout sheet -> (columns, rows). 3 is the layout that carries
# ruled lines, which is what PowerPoint itself calls a handout with lines.
HANDOUT_GRIDS = {2: (1, 2), 3: (1, 3), 4: (2, 2), 6: (2, 3)}
HANDOUT_PAGE = (595.276, 841.890)  # A4 portrait
HANDOUT_MARGIN = 36.0
HANDOUT_GUTTER = 16.0
HANDOUT_LINE_STEP = 22.0
HANDOUT_SLIDE_SHARE = 0.46  # width taken by the slides when lines are drawn


def task_ppt_handout(job_dir: str, params: dict) -> dict:
    """Lay the converted slide pages out several to a sheet.

    Runs on <job_dir>/output.pdf and replaces it. A failure here leaves the
    converted slides in place and reports it, because losing a conversion
    that already succeeded over a layout step would be the wrong trade.
    """
    try:
        per_page = int(params.get("per_page") or 0)
    except (TypeError, ValueError):
        return {"handout": "skipped"}
    if per_page not in HANDOUT_GRIDS:
        return {"handout": "skipped"}

    out = os.path.join(job_dir, "output.pdf")
    try:
        from pypdf import PdfReader, PdfWriter, Transformation
        from pypdf.generic import DecodedStreamObject, NameObject

        reader = PdfReader(out)
        if not len(reader.pages):
            return {"handout": "skipped"}

        writer = PdfWriter()
        cols, rows = HANDOUT_GRIDS[per_page]
        page_w, page_h = HANDOUT_PAGE
        lines = per_page == 3
        usable_w = (page_w - 2 * HANDOUT_MARGIN) * (HANDOUT_SLIDE_SHARE if lines else 1.0)
        cell_w = (usable_w - HANDOUT_GUTTER * (cols - 1)) / cols
        cell_h = (page_h - 2 * HANDOUT_MARGIN - HANDOUT_GUTTER * (rows - 1)) / rows
        per_sheet = cols * rows

        for start in range(0, len(reader.pages), per_sheet):
            batch = reader.pages[start:start + per_sheet]
            sheet = writer.add_blank_page(width=page_w, height=page_h)

            placed = []
            for slot, src_page in enumerate(batch):
                box = src_page.mediabox
                sw, sh = float(box.width), float(box.height)
                if sw <= 0 or sh <= 0:
                    continue
                col, row = slot % cols, slot // cols
                cx = HANDOUT_MARGIN + col * (cell_w + HANDOUT_GUTTER)
                # Slot 0 is the top left, so rows count down from the top.
                cy = page_h - HANDOUT_MARGIN - (row + 1) * cell_h - row * HANDOUT_GUTTER
                scale = min(cell_w / sw, cell_h / sh)
                placed.append(
                    (src_page, scale,
                     cx + (cell_w - sw * scale) / 2,
                     cy + (cell_h - sh * scale) / 2,
                     sw * scale, sh * scale, cy, cell_h)
                )

            ops = ["q", "0.72 G", "0.5 w"]
            for _, _, x, y, w, h, _, _ in placed:
                # Inflated by a point, so the hairline is not hidden beneath a
                # slide whose own background is white.
                ops.append(f"{x - 1:.2f} {y - 1:.2f} {w + 2:.2f} {h + 2:.2f} re S")
            if lines:
                left = HANDOUT_MARGIN + usable_w + HANDOUT_GUTTER
                right = page_w - HANDOUT_MARGIN
                ops.append("0.82 G")
                for _, _, _, _, _, _, cy, ch in placed:
                    count = max(1, int(ch // HANDOUT_LINE_STEP))
                    top = cy + ch - (ch - (count - 1) * HANDOUT_LINE_STEP) / 2
                    for i in range(count):
                        ly = top - i * HANDOUT_LINE_STEP
                        ops.append(f"{left:.2f} {ly:.2f} m {right:.2f} {ly:.2f} l S")
            ops.append("Q")

            # The decoration goes on first and pypdf merges the slides over
            # it. Writing a second content stream by hand and patching the
            # /Contents array puts a resolved object where a reference
            # belongs, and readers report the result as damaged.
            stream = DecodedStreamObject()
            stream.set_data(("\\n".join(ops)).encode("ascii"))
            sheet[NameObject("/Contents")] = writer._add_object(stream)

            for src_page, scale, x, y, *_ in placed:
                sheet.merge_transformed_page(src_page, Transformation().scale(scale).translate(x, y))

        tmp = os.path.join(job_dir, "handout.pdf")
        with open(tmp, "wb") as handle:
            writer.write(handle)
        if os.path.getsize(tmp) == 0:
            return {"handout": "skipped"}
        os.replace(tmp, out)
        return {"handout": "done"}
    except MemoryError:
        raise
    except Exception:
        return {"handout": "skipped"}


'''

WORKER_MAP_OLD = '    "excel_prepare": task_excel_prepare,\n    "merge": task_merge,\n}'
WORKER_MAP_NEW = '    "excel_prepare": task_excel_prepare,\n    "ppt_handout": task_ppt_handout,\n    "merge": task_merge,\n}'


# --------------------------------------------------------------------------
# 4. The export filter builders
# --------------------------------------------------------------------------

PDFA_ANCHOR = """PDFA_CONVERT_TO = (
    'pdf:writer_pdf_Export:{"SelectPdfVersion":{"type":"long","value":2}}'
)"""

FILTERS_NEW = '''PDFA_CONVERT_TO = (
    'pdf:writer_pdf_Export:{"SelectPdfVersion":{"type":"long","value":2}}'
)

# The same thing through Calc's export filter, for Excel to PDF.
CALC_PDFA_CONVERT_TO = (
    'pdf:calc_pdf_Export:{"SelectPdfVersion":{"type":"long","value":2}}'
)

SLIDE_RANGE_RE = re.compile(r"^[0-9]+(-[0-9]+)?(,[0-9]+(-[0-9]+)?)*$")


def clean_slide_range(text: str) -> str:
    """A page range LibreOffice will accept, or nothing at all.

    Anything that is not plain numbers, hyphens and commas is dropped rather
    than passed on, so a malformed box converts the whole deck instead of
    failing the job.
    """
    squeezed = re.sub(r"\\s+", "", str(text or ""))
    if not squeezed or len(squeezed) > 200:
        return ""
    return squeezed if SLIDE_RANGE_RE.match(squeezed) else ""


def impress_filter(content: str = "slides", hidden: str = "", slides: str = "",
                   quality: str = "", pdfa: bool = False) -> str:
    """Build the --convert-to value for a presentation.

    Returns plain "pdf" when nothing is asked for, so a request with no
    options is byte for byte the command this endpoint has always run.
    Every option here was checked against a real deck before it was added.
    """
    import json as _json

    data: dict = {}
    if content == "notes":
        data["ExportNotesPages"] = {"type": "boolean", "value": "true"}
    elif content == "only-notes":
        data["ExportNotesPages"] = {"type": "boolean", "value": "true"}
        data["ExportOnlyNotesPages"] = {"type": "boolean", "value": "true"}
    if office_flag(hidden):
        data["ExportHiddenSlides"] = {"type": "boolean", "value": "true"}
    page_range = clean_slide_range(slides)
    if page_range:
        data["PageRange"] = {"type": "string", "value": page_range}
    if str(quality or "").strip().lower() == "small":
        data["ReduceImageResolution"] = {"type": "boolean", "value": "true"}
        data["MaxImageResolution"] = {"type": "long", "value": 150}
        data["Quality"] = {"type": "long", "value": 70}
    if pdfa:
        data["SelectPdfVersion"] = {"type": "long", "value": 2}
    if not data:
        return "pdf"
    return "pdf:impress_pdf_Export:" + _json.dumps(data, separators=(",", ":"))'''


# --------------------------------------------------------------------------
# 5. The two endpoints
# --------------------------------------------------------------------------

ENDPOINTS_OLD = '''@app.post("/api/excel-to-pdf")
def excel_to_pdf(file: UploadFile = File(...)):
    with Job("/api/excel-to-pdf", file) as job:
        ext = safe_extension(file.filename, ".xlsx")
        job.start(ext)
        prepared = job.run_worker("excel_prepare", {"ext": ext})
        target = prepared.get("target") or f"input{ext}"
        if not re.fullmatch(r"(input\\.[a-z0-9]{1,5}|prepared\\.xlsx)", target):
            target = f"input{ext}"
        job.libreoffice_to_pdf(target)
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_ExcelToPDF_{job.base}.pdf")


@app.post("/api/ppt-to-pdf")
def ppt_to_pdf(file: UploadFile = File(...)):
    with Job("/api/ppt-to-pdf", file) as job:
        ext = safe_extension(file.filename, ".pptx")
        job.start(ext)
        job.libreoffice_to_pdf(f"input{ext}")
        return job.file_response("output.pdf", PDF_TYPE, f"OptimaxPDF_PPTtoPDF_{job.base}.pdf")'''

ENDPOINTS_NEW = '''# Spreadsheet formats openpyxl cannot open. They go through LibreOffice once
# to become .xlsx, so the layout options apply to every upload rather than to
# some of them.
EXCEL_NEEDS_CONVERSION = (".xls", ".ods", ".fods", ".xlt", ".xltx", ".sxc")


@app.post("/api/excel-to-pdf")
def excel_to_pdf(
    file: UploadFile = File(...),
    orientation: str = Form("auto"),
    fit: str = Form("width"),
    scale: str = Form(""),
    paper: str = Form("a4"),
    margins: str = Form("normal"),
    gridlines: str = Form(""),
    headings: str = Form(""),
    repeat_header: str = Form(""),
    sheets: str = Form(""),
    pdfa: str = Form(""),
):
    with Job("/api/excel-to-pdf", file) as job:
        ext = safe_extension(file.filename, ".xlsx")
        job.start(ext)

        source = f"input{ext}"
        notes = ""
        if ext in EXCEL_NEEDS_CONVERSION:
            try:
                source = job.libreoffice_to_xlsx(source)
            except ApiError as err:
                # A timeout is a timeout. Anything else and the file is
                # converted directly, which is what happens today, with the
                # page told that the layout options did not apply.
                if err.code == "timeout":
                    raise
                source = f"input{ext}"
                notes = "options_unavailable"

        prepared = job.run_worker(
            "excel_prepare",
            {
                "ext": ext,
                "source": source,
                "orientation": orientation,
                "fit": fit,
                "scale": scale,
                "paper": paper,
                "margins": margins,
                "gridlines": gridlines,
                "headings": headings,
                "repeat_header": repeat_header,
                "sheets": sheets,
            },
        )
        target = prepared.get("target") or source
        if not re.fullmatch(r"(input\\.[a-z0-9]{1,5}|prepared\\.xlsx|converted\\.xlsx)", target):
            target = f"input{ext}"

        if office_flag(pdfa):
            try:
                job.libreoffice_to_pdf(target, CALC_PDFA_CONVERT_TO)
            except ApiError as err:
                if err.code == "timeout":
                    raise
                job.libreoffice_to_pdf(target)
                notes = (notes + " pdfa_unavailable").strip()
        else:
            job.libreoffice_to_pdf(target)

        headers = {"X-Optimax-Notes": notes} if notes else None
        return job.file_response(
            "output.pdf", PDF_TYPE, f"OptimaxPDF_ExcelToPDF_{job.base}.pdf", headers=headers
        )


@app.post("/api/ppt-to-pdf")
def ppt_to_pdf(
    file: UploadFile = File(...),
    content: str = Form("slides"),
    hidden: str = Form(""),
    slides: str = Form(""),
    handout: str = Form(""),
    quality: str = Form(""),
    pdfa: str = Form(""),
):
    with Job("/api/ppt-to-pdf", file) as job:
        ext = safe_extension(file.filename, ".pptx")
        job.start(ext)

        grid = str(handout or "").strip()
        if grid not in ("2", "3", "4", "6"):
            grid = ""
        want_pdfa = office_flag(pdfa)
        # A handout is rebuilt from the converted pages, and that rebuild
        # cannot carry PDF/A markers across, so the two are never asked of
        # LibreOffice together. The page says so before the upload.
        convert_to = impress_filter(
            content=content, hidden=hidden, slides=slides,
            quality=quality, pdfa=want_pdfa and not grid,
        )

        notes = ""
        try:
            job.libreoffice_to_pdf(f"input{ext}", convert_to)
        except ApiError as err:
            # An option this build will not take should not cost somebody
            # their conversion. A plain run has nothing to fall back to.
            if err.code == "timeout" or convert_to == "pdf":
                raise
            job.libreoffice_to_pdf(f"input{ext}")
            notes = "options_unavailable"

        if grid:
            result = job.run_worker("ppt_handout", {"per_page": int(grid)})
            if result.get("handout") != "done":
                notes = (notes + " handout_unavailable").strip()
            if want_pdfa:
                notes = (notes + " pdfa_unavailable").strip()

        headers = {"X-Optimax-Notes": notes} if notes else None
        return job.file_response(
            "output.pdf", PDF_TYPE, f"OptimaxPDF_PPTtoPDF_{job.base}.pdf", headers=headers
        )'''


STEPS = [
    ("the per tool timeouts", LIMITS_OLD, LIMITS_NEW),
    ("the LibreOffice helpers", LO_ANCHOR, LO_NEW_METHOD),
    ("the export filter builders", PDFA_ANCHOR, FILTERS_NEW),
    ("the worker task list", WORKER_MAP_OLD, WORKER_MAP_NEW),
    ("the two endpoints", ENDPOINTS_OLD, ENDPOINTS_NEW),
]


def replace_prepare(source: str) -> str:
    """Swap the whole of task_excel_prepare for the new one."""
    start_marker = "def task_excel_prepare(job_dir: str, params: dict) -> dict:"
    end_marker = "\ndef task_merge(job_dir: str, params: dict) -> dict:"
    if source.count(start_marker) != 1 or source.count(end_marker) != 1:
        raise LookupError("task_excel_prepare as it was before this patch")
    start = source.index(start_marker)
    end = source.index(end_marker)
    if end < start:
        raise LookupError("task_excel_prepare in the position this patch expects")
    return source[:start] + PREPARE_NEW + source[end + 1:]


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

    updated = original
    try:
        for label, old, new in STEPS:
            if updated.count(old) != 1:
                raise LookupError(label)
            updated = updated.replace(old, new, 1)
        updated = replace_prepare(updated)
    except LookupError as err:
        print("Nothing was changed. Could not find:", err)
        print("The file is not the one this patch was written against. Send me your current server.py.")
        return 1

    # The marker that makes a second run a no-op. A comment on the first line
    # works whatever the module docstring looks like, which a replacement
    # inside that docstring would not.
    updated = "# " + MARK + ": Excel and PowerPoint conversion options, 22 November 2026.\n" + updated

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
    print("  Excel to PDF   orientation, fit, scale, paper, margins, gridlines,")
    print("                 row and column headings, repeat header row, sheet")
    print("                 picker, PDF/A, and .xls and .ods through a first pass")
    print("  PowerPoint     speaker notes, notes only, hidden slides, slide range,")
    print("                 handouts at 2, 3, 4 or 6 a page, smaller file, PDF/A")
    print("  Both           timeout raised from 120 to 180 seconds")
    print()
    print("Commit and push, wait for Render to finish, then check:")
    print("  curl -s https://optimaxpdf-docker-api.onrender.com/health")
    print()
    print("Then set OFFICE_OPTIONS_ENABLED to true in src/assets/js/config.js.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
