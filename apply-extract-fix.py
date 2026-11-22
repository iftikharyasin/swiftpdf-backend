#!/usr/bin/env python3
"""Two fixes for PDF to Excel, PowerPoint to PDF and Excel to PDF.

Run it from the folder holding server.py, after apply-extract-patch.py:

    python apply-extract-fix.py server.py

Then check it, which takes a second and needs nothing running:

    python check-server.py server.py

FIX 1, and this one is urgent.

apply-extract-patch.py replaced everything between `def task_pdf_to_excel`
and `def task_excel_prepare`, on the assumption that only the first function
sat in that span. It did not. The earlier office patch had placed thirteen
definitions there, and replacing the span deleted all of them:

    office_flag, _column_letter, excel_sheet_width_points,
    excel_clean_options, and the nine EXCEL_* constants

The file still compiled, because a missing name is a runtime error, not a
syntax error, so every check passed and the deploy went green. But
/api/excel-to-pdf and /api/ppt-to-pdf both call office_flag on every
request, so both have been answering 500 ever since, and task_excel_prepare
has been falling back to the untouched workbook, which means the Excel
layout options have been silently doing nothing.

This restores those definitions exactly as the office patch wrote them.

FIX 2.

pdfplumber returns rows of different lengths whenever a table is inferred
from spacing rather than ruled lines, which is most tables. Writing one of
those with a header row raised

    ValueError: 8 columns passed, passed data had 11 columns

so PDF to Excel failed on a real three page report whose first row held 8
cells and whose later rows held 11. Every table used in testing had ruled
lines, so every row was the same length and the case never arose. Tables are
now squared off before they are written.
"""

import os
import shutil
import sys

MARK = "OPTIMAX_EXTRACT_FIX_2"

# Exactly as the office patch wrote them, so the restore is a restore and
# not a rewrite.
RESTORE_BLOCK = r'''
def office_flag(value) -> bool:
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
'''

ANCHOR = "def task_excel_prepare(job_dir: str, params: dict) -> dict:"

OLD_XLSX = '''def _write_tables_xlsx(path, groups, header_row):
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
            width = max((len(row) for row in rows), default=0)'''

NEW_XLSX = '''def _rectangular(rows):
    \"\"\"Pad every row to the widest one, and report that width.

    pdfplumber gives rows of different lengths whenever a table is inferred
    from spacing rather than ruled lines, which is most of them. Handing a
    ragged table to pandas together with a header raises, because the column
    names and the data disagree on how many columns there are, and the
    visitor sees "we could not process your file" for a table that was read
    correctly. Squaring it off first costs nothing and removes the whole
    class of failure.
    \"\"\"
    width = max((len(row) for row in rows), default=0)
    return [list(row) + [""] * (width - len(row)) for row in rows], width


def _write_tables_xlsx(path, groups, header_row):
    import pandas as pd
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        used = set()
        for name, rows in groups:
            if not rows:
                continue
            rows, width = _rectangular(rows)
            if not width:
                continue
            sheet_name = safe_sheet_name(name, used)
            as_header = bool(header_row) and len(rows) > 1
            if as_header:
                frame = pd.DataFrame(rows[1:], columns=_unique_header(rows[0]))
                frame.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            sheet = writer.sheets[sheet_name]'''

OLD_CSV_MANY = '''        for name, rows in groups:
            if not rows:
                continue
            buffer = []'''

NEW_CSV_MANY = '''        for name, rows in groups:
            if not rows:
                continue
            rows, _ = _rectangular(rows)
            buffer = []'''

OLD_CSV_ONE = '''        with open(out, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            for row in (groups[0][1] if groups else []):
                writer.writerow(row)'''

NEW_CSV_ONE = '''        rows, _ = _rectangular(groups[0][1]) if groups and groups[0][1] else ([], 0)
        with open(out, "w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            for row in rows:
                writer.writerow(row)'''

OLD_END = '''    _require_output(os.path.join(job_dir, name))
    return {"file": name, "tables": len(found)}'''

NEW_END = '''    out_path = os.path.join(job_dir, name)
    if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
        # Nothing could be written, which means nothing usable was found.
        # "No tables in this PDF" is the truth; a generic failure is not.
        raise ApiError(400, "no_tables")
    _require_output(out_path)
    return {"file": name, "tables": len(found)}'''


def main() -> int:
    path = sys.argv[1] if len(sys.argv) > 1 else "server.py"
    if not os.path.exists(path):
        print(f"Nothing was changed. {path} is not here. Run this from the folder holding server.py.")
        return 1

    with open(path, encoding="utf-8") as handle:
        original = handle.read()

    if MARK in original:
        print("Nothing was changed. This fix is already in server.py.")
        return 0
    if "OPTIMAX_EXTRACT_V1" not in original:
        print("Nothing was changed. Apply apply-extract-patch.py first; this fixes that patch.")
        return 1

    updated = original
    restored = False

    # Fix 1: put back what the extract patch deleted.
    if "def office_flag(" not in updated:
        if updated.count(ANCHOR) != 1:
            print("Nothing was changed. Could not find: task_excel_prepare, where the restore goes.")
            return 1
        updated = updated.replace(ANCHOR, RESTORE_BLOCK.strip() + "\n\n\n" + ANCHOR, 1)
        restored = True

    # Fix 2: square off ragged tables.
    steps = [
        ("the spreadsheet writer", OLD_XLSX, NEW_XLSX),
        ("the CSV bundle writer", OLD_CSV_MANY, NEW_CSV_MANY),
        ("the single CSV writer", OLD_CSV_ONE, NEW_CSV_ONE),
        ("the end of task_pdf_to_excel", OLD_END, NEW_END),
    ]
    try:
        for label, old, new in steps:
            if updated.count(old) != 1:
                raise LookupError(label)
            updated = updated.replace(old, new, 1)
    except LookupError as err:
        print("Nothing was changed. Could not find:", err)
        print("Send me your current server.py and I will re-cut this.")
        return 1

    updated = "# " + MARK + ": restored helpers, ragged tables squared off, 22 November 2026.\n" + updated

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
    if restored:
        print("  RESTORED office_flag, _column_letter, excel_sheet_width_points,")
        print("  excel_clean_options and the EXCEL_* constants, which the extract")
        print("  patch had deleted. Excel to PDF and PowerPoint to PDF have been")
        print("  answering 500 on every request without them.")
    else:
        print("  The deleted helpers were already present, so nothing was restored.")
    print()
    print("  Ragged tables are squared off before writing, so the header row")
    print("  option no longer fails on a table without ruled lines.")
    print()
    print("Now run:  python check-server.py server.py")
    print("Then commit, push, and try Excel to PDF and PDF to Excel again.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
