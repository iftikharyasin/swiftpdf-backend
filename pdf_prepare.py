"""Clean a PDF before pdf2docx sees it.

Optimax PDF, PDF to Word. Three jobs, in this order:

1. Take out text that was never meant to be read. A publisher writes tags
   beside each equation for its own tooling, usually in text render mode 3,
   which draws nothing. pdf2docx cannot know that and hands them over as
   words, which is how "display-eq-Equ1" and "inline-eq-Eq" ended up in the
   10 November Springer conversion, 62 times.

2. Turn each display equation into a picture. A maths font with a broken
   ToUnicode map gives up nonsense, which is where "ββ ββ ββ" and "LLLLIIIII"
   came from. A picture of the equation is honest and readable, and it stops
   pdf2docx reading the aligned glyphs as a table. The equation number sits in
   the same region, so it is kept inside the picture.

3. Say whether the file has a two column page, which decides whether the
   caller may force one section per page.

Nothing here is allowed to lose a file. `prepare()` catches everything and
reports `skipped`, and the caller then converts the original.

The decisions live in pure functions near the top, so they can be checked
without a PDF: `is_math_font`, `is_eq_token`, `looks_like_math`, `group_rows`,
`equation_regions`, `column_count` and `drop_overlapping`. Only `prepare()`
touches PyMuPDF.
"""

from __future__ import annotations

import os
import re

# The publisher tags. One definition, shared with docx_repair, which is the
# safety net for a file that skipped this step. They were two copies until
# 10 November and they drifted apart, which is how 13 tags reached the Word
# file as text. Both modules ship together, so the import is safe.
from docx_repair import EQ_TOKEN

__all__ = ["prepare", "is_math_font", "is_eq_token", "looks_like_math",
           "group_rows", "equation_regions", "column_count", "drop_overlapping"]


# ---------------------------------------------------------------------------
# Tunables. Every one of these is a judgement that wants checking against real
# files, so they are named and gathered here rather than buried in the code.
# ---------------------------------------------------------------------------

MATH_SCORE_MIN = 0.50       # above this a row is treated as an equation
PLAIN_WORDS_VETO = 5        # this many ordinary words and it is a sentence
REGION_EDGE_GAP = 2.0       # points kept clear of the line above and below
# Three kinds of evidence, weighted. The font says what the glyphs came from,
# the characters say what they are, and the token lengths say how the line is
# built: an equation is mostly one and two character pieces, prose is not.
W_FONT = 0.45
W_CHARS = 0.30
W_TOKENS = 0.25
ROW_TOLERANCE = 3.0         # points, for grouping spans into one line
REGION_GAP = 14.0           # points, two equation rows closer than this merge
MIN_REGION_HEIGHT = 6.0     # points, ignore anything thinner
MAX_REGION_SHARE = 0.55     # a region taller than this share of the page is
                            # not an equation, it is the page
RENDER_DPI = 200
GUTTER_MIN_SHARE = 0.03     # a real gutter is about 25 pt on a 595 pt page
GUTTER_EDGE_SHARE = 0.25    # and this share of lines must stop and start at it
EDGE_MARGIN = 0.08          # ignore this much of each side when finding a gutter


# Font name fragments that mean mathematics or symbols rather than prose.
MATH_FONT_HINTS = (
    "cmmi", "cmsy", "cmex", "msam", "msbm", "mathematicalpi", "mathpi",
    "universalgreek", "universalmath", "greekwithmath", "symbol", "mtsy",
    "mtmi", "mtex", "euclid", "stixmath", "cambriamath", "lmmath", "xits",
    "asana", "neoeuler", "mathjax", "mathtime", "eufm", "eusm", "rsfs",
    "wasy", "stmary", "dsrom", "bbold",
)
# A subset prefix such as "BVCJSW+" tells us nothing, so it comes off first.
SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")

DISPLAY_EQ_TOKEN = re.compile(r"\bdisplay[-_]?eq", re.IGNORECASE)

# Greek, maths operators, arrows, set and logic symbols, sub and superscripts.
_MATH_RANGES = (
    (0x0370, 0x03FF), (0x1F00, 0x1FFF),      # Greek
    (0x2070, 0x209F), (0x2190, 0x21FF),      # super and subscripts, arrows
    (0x2200, 0x22FF), (0x2A00, 0x2AFF),      # maths operators
    (0x27C0, 0x27EF), (0x2980, 0x29FF),
    (0x1D400, 0x1D7FF),                      # maths alphanumerics
    (0x221A, 0x221A), (0x00B1, 0x00B1),
)

EQUATION_NUMBER = re.compile(r"^\(\s*\d+[a-z]?\s*\)$")


def _bare(name: str) -> str:
    return SUBSET_PREFIX.sub("", str(name or "")).lower().replace(" ", "").replace("-", "")


def is_math_font(name: str) -> bool:
    """True when a font name says mathematics or symbols."""
    bare = _bare(name)
    return any(hint in bare for hint in MATH_FONT_HINTS)


def is_eq_token(text: str) -> bool:
    """True when the text is, or contains, a publisher equation tag."""
    return bool(EQ_TOKEN.search(str(text or "")))


def _math_char_share(text: str) -> float:
    body = [ch for ch in str(text or "") if not ch.isspace()]
    if not body:
        return 0.0
    hits = 0
    for ch in body:
        point = ord(ch)
        if any(low <= point <= high for low, high in _MATH_RANGES):
            hits += 1
        elif ch in "=+<>≤≥≈∝×÷·^_/|∑∏∫∂∇":
            hits += 1
    return hits / len(body)


def _plain_words(text: str) -> int:
    """How many ordinary words of four letters or more a line holds.

    A word is ordinary when it is letters once the maths characters are taken
    out, so "elasticities" counts and "β1" does not.
    """
    count = 0
    for token in str(text or "").split():
        letters = re.sub(r"[^A-Za-z]", "", token)
        if len(letters) >= 4 and _math_char_share(token) < 0.34:
            count += 1
    return count


def _repeat_share(text: str) -> float:
    """Share of characters that simply repeat the one before them.

    A maths font read through a broken ToUnicode map produces runs like
    "ββ ββ" and "LLLLIIIII". Prose almost never does.
    """
    body = [ch for ch in str(text or "") if not ch.isspace()]
    if len(body) < 4:
        return 0.0
    repeats = sum(1 for index in range(1, len(body)) if body[index] == body[index - 1])
    return repeats / (len(body) - 1)


def looks_like_math(text: str, math_font_share: float) -> tuple[bool, float]:
    """Decide whether one line is a display equation. Returns (verdict, score)."""
    text = str(text or "")
    # Only a display tag marks a whole line as an equation. An inline tag sits
    # inside a line of prose, and that line must stay editable text.
    if DISPLAY_EQ_TOKEN.search(text):
        return True, 1.0
    text = EQ_TOKEN.sub(" ", text)
    if EQUATION_NUMBER.match(text.strip()):
        # An equation number on its own belongs to the equation above it, and
        # is picked up when regions are merged, not on its own account.
        return False, 0.0
    stripped = text.strip()
    if len(stripped) < 2:
        return False, 0.0
    tokens = stripped.split()
    # A line carrying this many ordinary words is a sentence, whatever its
    # fonts and symbols say. On 10 November the line "β1, β2, β3 are
    # elasticities and εit captures stochastic error..." was swept into an
    # equation picture; that cannot happen now.
    if _plain_words(stripped) >= PLAIN_WORDS_VETO:
        return False, 0.0
    short_share = (sum(1 for token in tokens if len(token) <= 3) / len(tokens)) if tokens else 0.0
    score = (W_FONT * float(math_font_share)
             + W_CHARS * _math_char_share(text)
             + W_TOKENS * short_share)
    # Heavy glyph repetition is strong evidence on its own.
    if _repeat_share(text) >= 0.45:
        score = max(score, 0.75)
    # A long line with ordinary word spacing is prose, whatever its fonts say.
    words = [word for word in stripped.split() if len(word) > 3]
    if len(stripped) > 60 and len(words) >= 6:
        score *= 0.5
    return score >= MATH_SCORE_MIN, score


def clean_math_text(text: str) -> str:
    """Tidy maths glyphs read straight out of a PDF.

    Two things go wrong. A maths font is often mapped to the Mathematical
    Alphanumeric block, so a beta arrives as U+1D6FD rather than U+03B2, and a
    broken ToUnicode map repeats every glyph, which is where "𝛽𝛽" came from.
    Compatibility normalisation fixes the first and halving a doubled run
    fixes the second.
    """
    import unicodedata

    text = unicodedata.normalize("NFKC", str(text or ""))
    out = []
    index = 0
    while index < len(text):
        char = text[index]
        run = 1
        while index + run < len(text) and text[index + run] == char:
            run += 1
        # A doubled glyph is the ToUnicode fault. Three or more in a row is
        # more likely to be real, such as an ellipsis of dots, so it is left.
        out.append(char * (run // 2 if run == 2 and not char.isspace() else run))
        index += run
    return re.sub(r"\s+", " ", "".join(out)).strip()


def inline_math_for(tag_box, visible_spans, pad: float = 1.0) -> str:
    """The visible maths sitting under an inline tag, cleaned.

    The tag is drawn on top of the glyphs it stands for, which is why its box
    cannot simply be covered: covering it would take the beta with it. Reading
    what is underneath lets the tag be replaced in the Word file instead.
    """
    hits = []
    for span in visible_spans:
        box = span.get("bbox")
        text = str(span.get("text") or "")
        if not box or not text.strip() or is_eq_token(text):
            continue
        if (box[0] - pad < tag_box[2] and box[2] + pad > tag_box[0]
                and box[1] - pad < tag_box[3] and box[3] + pad > tag_box[1]):
            hits.append((box[0], text))
    hits.sort()
    return clean_math_text("".join(text for _, text in hits))


def group_rows(spans, tolerance: float = ROW_TOLERANCE):
    """Group spans into visual lines by the middle of their box.

    Each span is a dict with `bbox` as (x0, y0, x1, y1), `text`, `font`.
    Returns rows sorted top to bottom, each a list of spans left to right.
    """
    items = [span for span in spans if span.get("bbox")]
    items.sort(key=lambda span: ((span["bbox"][1] + span["bbox"][3]) / 2.0, span["bbox"][0]))
    rows: list[list[dict]] = []
    for span in items:
        middle = (span["bbox"][1] + span["bbox"][3]) / 2.0
        if rows:
            last = rows[-1][0]["bbox"]
            if abs(middle - (last[1] + last[3]) / 2.0) <= tolerance:
                rows[-1].append(span)
                continue
        rows.append([span])
    for row in rows:
        row.sort(key=lambda span: span["bbox"][0])
    return rows


def _row_box(row):
    return (min(s["bbox"][0] for s in row), min(s["bbox"][1] for s in row),
            max(s["bbox"][2] for s in row), max(s["bbox"][3] for s in row))


def _row_text(row):
    return "".join(str(s.get("text") or "") for s in row)


def _row_math_font_share(row):
    total = sum(len(str(s.get("text") or "")) for s in row)
    if not total:
        return 0.0
    math = sum(len(str(s.get("text") or "")) for s in row if is_math_font(s.get("font")))
    return math / total


def equation_regions(rows, page_box, body_box=None):
    """Merge equation rows into regions to render, widest body width.

    `page_box` and `body_box` are (x0, y0, x1, y1). Returns a list of
    (x0, y0, x1, y1) in page coordinates, top to bottom.
    """
    page_height = max(1.0, page_box[3] - page_box[1])
    left = (body_box or page_box)[0]
    right = (body_box or page_box)[2]

    flagged = []
    for row in rows:
        verdict, _ = looks_like_math(_row_text(row), _row_math_font_share(row))
        flagged.append(verdict)

    regions = []
    index = 0
    while index < len(rows):
        if not flagged[index]:
            index += 1
            continue
        box = list(_row_box(rows[index]))
        last = index
        probe = index + 1
        while probe < len(rows):
            next_box = _row_box(rows[probe])
            close = next_box[1] - box[3] <= REGION_GAP
            if not close:
                break
            text = _row_text(rows[probe]).strip()
            # An equation number, or another equation row, joins the region.
            if flagged[probe] or EQUATION_NUMBER.match(text):
                box[0] = min(box[0], next_box[0])
                box[1] = min(box[1], next_box[1])
                box[2] = max(box[2], next_box[2])
                box[3] = max(box[3], next_box[3])
                last = probe
                probe += 1
                continue
            break
        # Keep clear of the lines above and below. The picture spans the body
        # width, so without this the 2 pt of air can reach far enough into a
        # neighbouring line to put it in the image and then redact it away.
        top = box[1] - REGION_EDGE_GAP
        bottom = box[3] + REGION_EDGE_GAP
        if index > 0:
            above = _row_box(rows[index - 1])[3]
            top = max(top, (above + box[1]) / 2.0 if above < box[1] else top)
        if last + 1 < len(rows):
            below = _row_box(rows[last + 1])[1]
            bottom = min(bottom, (box[3] + below) / 2.0 if below > box[3] else bottom)

        height = bottom - top
        if MIN_REGION_HEIGHT <= height <= page_height * MAX_REGION_SHARE:
            # Widen to the body column so the picture is not cropped tight.
            regions.append((min(left, box[0]) - 2.0, top,
                            max(right, box[2]) + 2.0, bottom))
        index = last + 1
    return regions


def column_count(boxes, page_box):
    """1 or 2, from the gap in the middle of the page.

    A two column page has a vertical band with no text in it. Looking for that
    band is cheaper and steadier than clustering the boxes.
    """
    if not boxes:
        return 1
    page_left, page_right = page_box[0], page_box[2]
    width = page_right - page_left
    if width <= 0:
        return 1
    inset = width * EDGE_MARGIN
    low, high = page_left + inset, page_right - inset
    # A line that spans most of the width is a title or a full width figure,
    # and tells us nothing about columns.
    narrow = [box for box in boxes if (box[2] - box[0]) < width * 0.75]
    if len(narrow) < 6:
        return 1

    steps = 120
    step = width / steps
    covered = [False] * steps
    for box in narrow:
        start = max(0, int((box[0] - page_left) / step))
        end = min(steps - 1, int((box[2] - page_left) / step))
        for slot in range(start, end + 1):
            covered[slot] = True

    # Longest empty run whose middle sits near the middle of the page.
    best = (0, 0, 0)  # length, first slot, last slot
    run_start = None
    for slot in range(steps + 1):
        inside = (slot < steps and not covered[slot]
                  and low <= page_left + (slot + 0.5) * step <= high)
        if inside:
            if run_start is None:
                run_start = slot
            continue
        if run_start is not None:
            length = slot - run_start
            if length > best[0]:
                best = (length, run_start, slot - 1)
            run_start = None
    if best[0] == 0:
        return 1

    gap_left = page_left + best[1] * step
    gap_right = page_left + (best[2] + 1) * step
    middle = (gap_left + gap_right) / 2.0
    if abs(middle - (page_left + width / 2.0)) > width * 0.18:
        return 1
    if (gap_right - gap_left) < width * GUTTER_MIN_SHARE:
        return 1

    # The decisive test: a gutter is where lines stop and where lines start.
    # A chance patch of white space in a single column fails this.
    tolerance = max(6.0, width * 0.03)
    ends = sum(1 for box in narrow if abs(box[2] - gap_left) <= tolerance)
    starts = sum(1 for box in narrow if abs(box[0] - gap_right) <= tolerance)
    needed = max(3, int(len(narrow) * GUTTER_EDGE_SHARE))
    return 2 if ends >= needed and starts >= needed else 1


def drop_overlapping(rects, keep_rects, pad: float = 0.5):
    """Keep only the rects that touch none of `keep_rects`.

    Invisible text is removed by covering its box, and covering a box takes
    out whatever else sits under it. A tag that overlaps real words is left
    alone rather than risked.
    """
    out = []
    for rect in rects:
        clash = False
        for keep in keep_rects:
            if (rect[0] - pad < keep[2] and rect[2] + pad > keep[0]
                    and rect[1] - pad < keep[3] and rect[3] + pad > keep[1]):
                clash = True
                break
        if not clash:
            out.append(rect)
    return out


# ---------------------------------------------------------------------------
# The only part that needs PyMuPDF
# ---------------------------------------------------------------------------

def _new_report() -> dict:
    return {
        "ok": False,
        "skipped": "",
        "pages": 0,
        "invisible_removed": 0,
        "tokens_removed": 0,
        "equations_imaged": 0,
        "two_column_pages": 0,
        "columns": 1,
        "inline_math": [],
    }


def _page_spans(page):
    """Visible text spans, and the boxes of invisible ones.

    `get_texttrace` is the only reading that carries the text render mode, and
    mode 3 and 7 draw nothing. Where it is not available the file is still
    handled, just without the invisible text pass.
    """
    visible, hidden = [], []
    trace = None
    try:
        trace = page.get_texttrace()
    except Exception:
        trace = None

    if trace:
        for span in trace:
            chars = span.get("chars") or []
            text = "".join(chr(ch[0]) for ch in chars if ch and ch[0])
            boxes = [ch[3] for ch in chars if ch and len(ch) > 3 and ch[3]]
            if boxes:
                box = (min(b[0] for b in boxes), min(b[1] for b in boxes),
                       max(b[2] for b in boxes), max(b[3] for b in boxes))
            else:
                box = tuple(span.get("bbox") or (0, 0, 0, 0))
            record = {"text": text, "font": span.get("font") or "", "bbox": box}
            if int(span.get("type", 0)) in (3, 7):
                hidden.append(record)
            else:
                visible.append(record)
        return visible, hidden

    # No trace available: read the ordinary way and treat nothing as hidden.
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                visible.append({"text": span.get("text") or "",
                                "font": span.get("font") or "",
                                "bbox": tuple(span.get("bbox") or (0, 0, 0, 0))})
    return visible, []


def _body_box(spans, page_box):
    if not spans:
        return page_box
    return (min(s["bbox"][0] for s in spans), page_box[1],
            max(s["bbox"][2] for s in spans), page_box[3])


def prepare(src: str, dst: str, budget=None, max_equations_per_page: int = 12) -> dict:
    """Write a cleaned copy of `src` to `dst`. Never raises.

    Returns a report. `ok` False means the caller should convert the original,
    and `skipped` says why.
    """
    report = _new_report()
    try:
        import fitz  # PyMuPDF, installed with pdf2docx
    except Exception:
        report["skipped"] = "no_pymupdf"
        return report

    doc = None
    try:
        doc = fitz.open(src)
        report["pages"] = doc.page_count
        two_column = 0

        for page in doc:
            if budget is not None and not budget.have(1.0):
                report["skipped"] = "out_of_time"
                break

            page_box = tuple(page.rect)
            visible, hidden = _page_spans(page)
            body = _body_box(visible, page_box)

            boxes = [s["bbox"] for s in visible if str(s.get("text") or "").strip()]
            columns_here = column_count(boxes, page_box)
            if columns_here == 2:
                two_column += 1

            # 1. Cover invisible text and the publisher tags, but never where
            #    they sit on top of words a reader can see.
            targets = [s["bbox"] for s in hidden]
            tokens = [s["bbox"] for s in visible if is_eq_token(s.get("text"))]
            keep = [s["bbox"] for s in visible
                    if str(s.get("text") or "").strip() and not is_eq_token(s.get("text"))]
            safe = drop_overlapping(targets + tokens, keep)
            report["invisible_removed"] += len(drop_overlapping(targets, keep))
            report["tokens_removed"] += len(drop_overlapping(tokens, keep))

            # A tag that sits on top of the glyphs it stands for cannot be
            # covered without taking them too, so it stays in the PDF and
            # reaches the Word file as text. Read what is underneath it now,
            # in reading order, so docx_repair can put that back in its place.
            covered = {tuple(rect) for rect in safe}
            stuck = [span for span in (hidden + visible)
                     if is_eq_token(span.get("text")) and tuple(span["bbox"]) not in covered]
            stuck.sort(key=lambda span: (round(span["bbox"][1], 1), span["bbox"][0]))
            for span in stuck:
                report["inline_math"].append({
                    "tag": str(span.get("text") or "").strip(),
                    "text": inline_math_for(span["bbox"], visible),
                })

            # 2. Find the display equations and render them before anything is
            #    taken away, because the picture has to show what was there.
            regions = []
            if columns_here == 2:
                # Each column on its own, so a picture never reaches into the
                # other column and erases a line of it.
                middle = (page_box[0] + page_box[2]) / 2.0
                for side in ("left", "right"):
                    part = [s for s in visible
                            if ((s["bbox"][0] + s["bbox"][2]) / 2.0 < middle) == (side == "left")]
                    if not part:
                        continue
                    if side == "left":
                        column_box = (body[0], body[1], min(body[2], middle - 1.0), body[3])
                    else:
                        column_box = (max(body[0], middle + 1.0), body[1], body[2], body[3])
                    for region in equation_regions(group_rows(part), page_box, column_box):
                        regions.append((max(region[0], column_box[0] - 2.0), region[1],
                                        min(region[2], column_box[2] + 2.0), region[3]))
                regions.sort(key=lambda region: region[1])
            else:
                regions = equation_regions(group_rows(visible), page_box, body)
            regions = regions[:max_equations_per_page]
            pictures = []
            for region in regions:
                try:
                    rect = fitz.Rect(*region) & page.rect
                    if rect.is_empty:
                        continue
                    pix = page.get_pixmap(clip=rect, dpi=RENDER_DPI)
                    pictures.append((rect, pix.tobytes("png")))
                except Exception:
                    continue

            for rect in safe + [rect for rect, _ in pictures]:
                try:
                    page.add_redact_annot(fitz.Rect(*rect) if not hasattr(rect, "x0") else rect)
                except Exception:
                    continue
            try:
                page.apply_redactions(
                    images=getattr(fitz, "PDF_REDACT_IMAGE_NONE", 0),
                    graphics=getattr(fitz, "PDF_REDACT_LINE_ART_NONE", 0),
                )
            except TypeError:
                # An older PyMuPDF takes no options.
                page.apply_redactions()
            except Exception:
                pass

            for rect, png in pictures:
                try:
                    page.insert_image(rect, stream=png)
                    report["equations_imaged"] += 1
                except Exception:
                    continue

        report["two_column_pages"] = two_column
        report["columns"] = 2 if two_column else 1
        doc.save(dst, garbage=3, deflate=True)
        report["ok"] = os.path.exists(dst) and os.path.getsize(dst) > 0
        if not report["ok"]:
            report["skipped"] = report["skipped"] or "empty_output"
    except Exception as exc:
        report["ok"] = False
        report["skipped"] = report["skipped"] or type(exc).__name__
    finally:
        try:
            if doc is not None:
                doc.close()
        except Exception:
            pass
    return report
