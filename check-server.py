#!/usr/bin/env python3
"""Smoke test the patched server.py before you push it.

    python check-server.py server.py

Why this exists. The frontend has 536 tests that run on every build, and the
backend had none. On 22 November PDF to Excel went live and failed on a real
file: pdfplumber returns rows of different lengths whenever a table is
inferred from spacing rather than ruled lines, and writing one of those with
a header row raised

    ValueError: 8 columns passed, passed data had 11 columns

which the worker turned into a generic 500. Every table I had tested with
had ruled lines, so every row was the same length and the case never came up.

This script loads the pure functions out of server.py and runs them against
the shapes that actually occur, including the ragged one. It needs pandas,
openpyxl and pypdf, which the server already depends on. It does not start
the server, touch the network, or need Docker, so it runs in a second.

It is not part of the Render build, because the frontend's build image has no
Python. Run it yourself after applying a patch and before pushing.
"""

import ast
import os
import re
import sys
import tempfile

WANTED_FUNCS = {
    "_rectangular",
    "_unique_header",
    "_group_tables",
    "_write_tables_xlsx",
    "_write_tables_csv",
    "_CsvSink",
    "safe_sheet_name",
    "expand_page_range",
    "clean_ocr_languages",
    "clean_slide_range",
    "impress_filter",
    "office_flag",
    "excel_clean_options",
    "excel_sheet_width_points",
    "_column_letter",
}
WANTED_PREFIXES = ("MAX_PREVIEW_", "EXTRACT_", "HANDOUT_", "EXCEL_", "CALC_", "SLIDE_RANGE_RE")
WANTED_FUNCS |= {"office_flag"}

passed = 0
failed = []


def check(label, condition, detail=""):
    global passed
    if condition:
        passed += 1
    else:
        failed.append(f"{label}{(': ' + detail) if detail else ''}")


def try_check(label, fn):
    """Run a check that might raise, and record the raise as a failure.

    The first version of this script let an exception end the whole run, so
    one broken thing hid every check after it. A test tool that stops at the
    first problem tells you less than one that finishes.
    """
    try:
        check(label, bool(fn()))
    except Exception as err:
        check(label, False, f"{type(err).__name__}: {err}")


def load(path):
    """Pull the pure helpers out of server.py without importing FastAPI."""
    tree = ast.parse(open(path, encoding="utf-8").read())
    keep = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.ClassDef)) and node.name in WANTED_FUNCS:
            keep.append(node)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id.startswith(WANTED_PREFIXES):
                    keep.append(node)
                    break
    namespace = {"os": os, "re": re, "__name__": "checked"}
    exec(compile(ast.fix_missing_locations(ast.Module(body=keep, type_ignores=[])), path, "exec"), namespace)
    return namespace


def undefined_names(path):
    """Every global name the file uses but never defines.

    This is the check that matters most. On 22 November a patch replaced a
    span of server.py that turned out to hold thirteen definitions besides
    the one it meant to replace. Python only complains about a missing name
    when the line runs, so the file compiled, every build check passed, the
    deploy went green, and two endpoints answered 500 on every request. A
    missing name is found in milliseconds by looking, and never by running
    a build.
    """
    import builtins

    tree = ast.parse(open(path, encoding="utf-8").read())
    defined = set(dir(builtins))
    # Module globals Python provides, which no file declares.
    defined.update({"__file__", "__name__", "__doc__", "__package__", "__spec__", "__loader__", "__builtins__", "__debug__"})
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            # Arguments are picked up by the ast.arg branch below, which
            # catches every one of them wherever it appears. Reaching into
            # node.args here breaks on a class, which has none.
            defined.add(node.name)
        elif isinstance(node, ast.Name) and isinstance(node.ctx, (ast.Store, ast.Del)):
            defined.add(node.id)
        elif isinstance(node, (ast.Import, ast.ImportFrom)):
            for alias in node.names:
                defined.add((alias.asname or alias.name).split(".")[0])
        elif isinstance(node, ast.ExceptHandler) and node.name:
            defined.add(node.name)
        elif isinstance(node, (ast.comprehension,)):
            for small in ast.walk(node.target):
                if isinstance(small, ast.Name):
                    defined.add(small.id)
        elif isinstance(node, ast.arg):
            defined.add(node.arg)
        elif isinstance(node, (ast.With, ast.AsyncWith)):
            for item in node.items:
                if item.optional_vars:
                    for small in ast.walk(item.optional_vars):
                        if isinstance(small, ast.Name):
                            defined.add(small.id)
        elif isinstance(node, ast.Global):
            defined.update(node.names)

    used = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load) and node.id not in defined:
            used.setdefault(node.id, node.lineno)
    return sorted(used.items(), key=lambda pair: pair[1])


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "server.py"
    if not os.path.exists(path):
        print(f"{path} is not here. Run this from the folder holding server.py.")
        return 1

    for name, line in undefined_names(path):
        failed.append(f"name used but never defined: {name} (first used on line {line})")
    check("every name the file uses is defined somewhere", not failed)

    ns = load(path)

    missing = [name for name in ("_group_tables", "_write_tables_xlsx") if name not in ns]
    if missing:
        print(f"server.py does not hold {missing}. Apply the extract patch first.")
        return 1

    tmp = tempfile.mkdtemp(prefix="optimax-check-")

    # ---- The bug that reached production on 22 November ----
    if "_rectangular" not in ns:
        failed.append("ragged tables: _rectangular is missing, so apply-extract-fix.py has not been applied")
    else:
        rows, width = ns["_rectangular"]([["a", "b"], ["c", "d", "e", "f"]])
        check("ragged: padded to the widest row", width == 4 and all(len(r) == 4 for r in rows))
        check("ragged: nothing is lost", rows[0][:2] == ["a", "b"] and rows[1] == ["c", "d", "e", "f"])
        check("ragged: an empty table is handled", ns["_rectangular"]([]) == ([], 0))

    ragged = [
        ["one", "two", "three", "four", "five", "six", "seven", "eight"],
        ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k"],
        ["l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v"],
    ]
    shapes = {
        "ragged, short header": ragged,
        "ragged, long header": [["a"] * 11, ["b"] * 3, ["c"] * 7],
        "one row only": [["only", "row"]],
        "an empty row in the middle": [["a", "b"], [], ["c", "d", "e"]],
        "one cell": [["x"]],
    }
    for label, rows in shapes.items():
        for header in (False, True):
            name = f"{label}, header={header}"
            try:
                out = os.path.join(tmp, f"{abs(hash(name))}.xlsx")
                ns["_write_tables_xlsx"](out, [("Table_Data", rows)], header)
                check(f"xlsx: {name}", os.path.exists(out) and os.path.getsize(out) > 0, "no file written")
            except Exception as err:
                check(f"xlsx: {name}", False, f"{type(err).__name__}: {err}")

    # ---- CSV, both shapes ----
    try:
        single = ns["_write_tables_csv"](tmp, [("Only", ragged)])
        check("csv: one group gives one file", single == "output.csv")
        many = ns["_write_tables_csv"](tmp, [("A", ragged), ("B", [["x"], ["y", "z"]])])
        check("csv: several groups give a zip", many == "output.zip")
    except Exception as err:
        check("csv", False, f"{type(err).__name__}: {err}")

    # ---- Sheet layouts ----
    found = [(1, [["a", "b"]]), (1, [["c", "d"]]), (2, [["e", "f"]])]
    check("sheets: one keeps the old sheet name", ns["_group_tables"](found, "one")[0][0] == "Table_Data")
    check("sheets: per table gives one each", len(ns["_group_tables"](found, "table")) == 3)
    check("sheets: per page groups by page", [n for n, _ in ns["_group_tables"](found, "page")] == ["Page 1", "Page 2"])

    # ---- Names Excel will accept ----
    used = set()
    check("sheet names: illegal characters go", "/" not in ns["safe_sheet_name"]("A/B:C*D", used))
    check("sheet names: a repeat is made unique", ns["safe_sheet_name"]("A B C D", used) != "A B C D")
    check("sheet names: 31 characters at most", len(ns["safe_sheet_name"]("x" * 60, used)) <= 31)

    # ---- Page ranges and languages, which take text from the browser ----
    check("range: a normal one", ns["expand_page_range"]("1-3,7", 10) == [1, 2, 3, 7])
    check("range: trimmed to the document", ns["expand_page_range"]("1-99", 3) == [1, 2, 3])
    check("range: backwards is read either way", ns["expand_page_range"]("9-3", 10) == list(range(3, 10)))
    check("range: rubbish means every page", ns["expand_page_range"]("1;drop table", 10) is None)
    check("range: past the end means every page", ns["expand_page_range"]("99", 2) is None)
    check("languages: a real pair", ns["clean_ocr_languages"]("urd+ara") == "urd+ara")
    check("languages: rubbish falls back to English", ns["clean_ocr_languages"]("'; rm -rf /") == "eng")
    check("languages: no more than three", len(ns["clean_ocr_languages"]("a+b+c+d".replace("a", "urd").replace("b", "ara").replace("c", "hin").replace("d", "spa")).split("+")) == 3)

    # ---- The PowerPoint export filter, if that patch is in ----
    if "impress_filter" in ns:
        try_check("impress: nothing asked for means a plain run", lambda: ns["impress_filter"]() == "pdf")
        try_check("impress: notes pages", lambda: "ExportNotesPages" in ns["impress_filter"](content="notes"))
        try_check("impress: a range is passed through", lambda: '"1-2"' in ns["impress_filter"](slides="1-2"))
        try_check("impress: a bad range is dropped", lambda: "PageRange" not in ns["impress_filter"](slides="1;2"))

    print()
    for line in failed:
        print(f"  FAILED  {line}")
    print()
    if failed:
        print(f"{passed} checks passed, {len(failed)} FAILED. Do not push this server.py.")
        return 1
    print(f"All {passed} checks passed. server.py looks good to push.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
