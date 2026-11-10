"""Repair a DOCX produced by pdf2docx.

Optimax PDF, PDF to Word. This module never talks to the network and never
raises on normal input: `repair()` returns a report, and the caller keeps the
unrepaired file if anything goes wrong.

What it fixes, with the fault number from the 9 November test file:

1. Page count grows
   pdf2docx writes one Word section per PDF page and leaves the header
   distance at Word's default 36 pt while setting the top margin to 16 pt,
   so the running head prints over the body and the body spills. Here the
   head moves into a real header, the margin grows by the height of that
   line, and in "flow" mode the per page sections become one document.

2. Overlapping columns and the head printing mid text
   pdf2docx writes <w:cols w:num="1"> but never clears the <w:col> children
   between sections, so section 2 of a file carries two of them, section 3
   three, and Word lays the page out in that many overlapping columns.
   fix_columns rewrites every <w:cols> to one column.

3. PostScript font names
   resolve_font drops the subset prefix and the merge index, strips the PSMT,
   MT and PS suffixes, and maps the result to a family that is installed on
   the reader's machine, so Word never substitutes silently.

4. Line end hyphens
   pdf2docx's own delete_end_line_hyphen handles these at the source, where
   the line ends are known. rejoin_hyphens is the safety net for whatever
   survives, and it uses the document's own vocabulary rather than a
   dictionary.

5. A logo drawn in a symbol font
   drop_marker_runs finds it by how little the font is used. Font flags are
   no help: the Springer mark in the test file is drawn in a font whose own
   /Flags call it a non symbolic serif.

6. Hard breaks instead of flowing paragraphs
   breaks_to_spaces and merge_broken_paragraphs, in "flow" mode.

Also, from the 9 November review:

a. An inline image clipped by its paragraph's exact line height
   fix_image_paragraphs.
b and c. Running heads that alternate between left and right pages
   lift_running_text sets a real odd page and even page header, so both are
   kept and each lands on the right page.

Modes
  "flow"   Editable text. One section, merged paragraphs, no page breaks.
  "layout" Keep layout. One section per page and the page breaks kept.
"""

from __future__ import annotations

import re
from collections import Counter

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

__all__ = ["repair", "new_report", "resolve_font"]


# ---------------------------------------------------------------------------
# Fault 3. Font names
# ---------------------------------------------------------------------------

SUBSET_PREFIX = re.compile(r"^[A-Z]{6}\+")
MERGE_INDEX = re.compile(r"(?<=[A-Za-z])\d+$")
PS_SUFFIX = re.compile(r"(PSMT|PSM|MT|PS)$")

STYLE_WORDS = ("BoldItalic", "BoldOblique", "SemiboldItalic", "Semibold",
               "Bold", "Italic", "Oblique", "Light", "Medium", "Black",
               "Regular", "Roman", "Book", "SemiCn", "Cond", "Condensed", "Cn")

# PostScript stem, lower case and without spaces, to a family that really is
# installed on Windows and macOS. Anything not listed is classified below.
FAMILY_MAP = {
    "timesnewroman": "Times New Roman",
    "times": "Times New Roman",
    "nimbusromno9l": "Times New Roman",
    "liberationserif": "Times New Roman",
    "thorndale": "Times New Roman",
    "arial": "Arial",
    "arialnarrow": "Arial Narrow",
    "helvetica": "Arial",
    "helveticaneue": "Arial",
    "nimbussanl": "Arial",
    "liberationsans": "Arial",
    "albany": "Arial",
    "couriernew": "Courier New",
    "courier": "Courier New",
    "liberationmono": "Courier New",
    "calibri": "Calibri",
    "carlito": "Calibri",
    "cambria": "Cambria",
    "caladea": "Cambria",
    "georgia": "Georgia",
    "garamond": "Garamond",
    "verdana": "Verdana",
    "tahoma": "Tahoma",
    "trebuchetms": "Trebuchet MS",
    "segoeui": "Segoe UI",
    "palatino": "Palatino Linotype",
    "palatinolinotype": "Palatino Linotype",
    "bookantiqua": "Book Antiqua",
    "centurygothic": "Century Gothic",
    "consolas": "Consolas",
    "symbol": "Symbol",
    "wingdings": "Wingdings",
    "zapfdingbats": "Wingdings",
    "cambriamath": "Cambria Math",
    "stixgeneral": "Cambria Math",
    "cmr": "Cambria Math",
    # Foundry and Adobe families a reader almost never has. Word would
    # substitute these without saying so, so we choose the substitute and
    # record it in the report instead.
    "myriadpro": "Calibri",
    "myriad": "Calibri",
    "minionpro": "Cambria",
    "minion": "Cambria",
    "sourcesanspro": "Calibri",
    "sourceserifpro": "Cambria",
    "frutiger": "Calibri",
    "univers": "Arial",
    "futura": "Century Gothic",
    "optima": "Calibri",
    "charter": "Georgia",
}

SERIF_HINTS = ("times", "serif", "roman", "georgia", "garamond", "minion",
               "book", "cambria", "palatino", "caslon", "baskerville",
               "slab", "antiqua", "schoolbook", "pro", "text")
MONO_HINTS = ("mono", "courier", "consol", "code", "typewriter", "pitch")

SAFE_FAMILIES = set(FAMILY_MAP.values())


def split_ps_name(name: str) -> tuple[str, str]:
    """Return (stem, style) for a PostScript font name."""
    clean = SUBSET_PREFIX.sub("", (name or "").strip())
    clean = MERGE_INDEX.sub("", clean)
    style = ""
    for separator in (",", "-"):
        if separator in clean:
            clean, _, style = clean.partition(separator)
            break
    clean = PS_SUFFIX.sub("", clean)
    style = PS_SUFFIX.sub("", style)
    if not style:
        for word in STYLE_WORDS:
            if clean.endswith(word) and len(clean) > len(word):
                clean, style = clean[: -len(word)], word
                break
    return clean, style


def _camel_to_words(stem: str) -> str:
    spaced = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", stem)
    spaced = re.sub(r"(?<=[A-Z])(?=[A-Z][a-z])", " ", spaced)
    return re.sub(r"\s+", " ", spaced).strip()


def resolve_font(name: str) -> tuple[str, bool, bool, bool]:
    """Map a PDF font name to (family, bold, italic, substituted)."""
    stem, style = split_ps_name(name)
    key = stem.lower().replace(" ", "")
    low_style = style.lower()
    bold = any(word in low_style for word in ("bold", "black", "semibold", "heavy"))
    italic = any(word in low_style for word in ("italic", "oblique"))

    if key in FAMILY_MAP:
        family = FAMILY_MAP[key]
        return family, bold, italic, family.lower().replace(" ", "") != key

    pretty = _camel_to_words(stem)
    if pretty in SAFE_FAMILIES:
        return pretty, bold, italic, False
    if any(hint in key for hint in MONO_HINTS):
        return "Courier New", bold, italic, True
    if any(hint in key for hint in SERIF_HINTS):
        return "Times New Roman", bold, italic, True
    return "Calibri", bold, italic, True


def _write_fonts(rpr, family: str) -> None:
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.insert(0, rfonts)
    for attribute in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attribute), family)


def _ensure_flag(rpr, tag: str, wanted: bool) -> None:
    if wanted and rpr.find(qn(tag)) is None:
        rpr.append(rpr.makeelement(qn(tag), {}))


def fix_fonts(roots, report) -> None:
    """Rewrite every rFonts in the body, the styles and the header parts."""
    for root in roots:
        if root is None:
            continue
        for rpr in root.iter(qn("w:rPr")):
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is None:
                continue
            original = (rfonts.get(qn("w:ascii")) or rfonts.get(qn("w:hAnsi"))
                        or rfonts.get(qn("w:cs")) or "")
            if not original:
                continue
            family, bold, italic, substituted = resolve_font(original)
            _write_fonts(rpr, family)
            _ensure_flag(rpr, "w:b", bold)
            _ensure_flag(rpr, "w:i", italic)
            report["fonts"][original] = {"family": family, "substituted": substituted}


# ---------------------------------------------------------------------------
# Fault 2. Columns
# ---------------------------------------------------------------------------

def fix_columns(doc, report) -> None:
    """One column per section, as wide as the text area."""
    for sect_pr in doc.element.body.iter(qn("w:sectPr")):
        cols = sect_pr.find(qn("w:cols"))
        if cols is None:
            continue
        children = cols.findall(qn("w:col"))
        if len(children) > 1:
            report["columns_fixed"] += 1
            report["column_counts"].append(len(children))

        width = None
        page_size = sect_pr.find(qn("w:pgSz"))
        margins = sect_pr.find(qn("w:pgMar"))
        if page_size is not None and margins is not None:
            try:
                width = (int(page_size.get(qn("w:w")))
                         - int(margins.get(qn("w:left")) or 0)
                         - int(margins.get(qn("w:right")) or 0))
            except (TypeError, ValueError):
                width = None

        for child in children:
            cols.remove(child)
        cols.set(qn("w:num"), "1")
        cols.set(qn("w:equalWidth"), "1")
        cols.set(qn("w:space"), "720")
        if width and width > 0:
            column = cols.makeelement(qn("w:col"), {})
            column.set(qn("w:w"), str(width))
            cols.append(column)


# ---------------------------------------------------------------------------
# Review point a. Images clipped by an exact line height
# ---------------------------------------------------------------------------

def has_drawing(paragraph_element) -> bool:
    return bool(paragraph_element.findall(".//" + qn("w:drawing"))
                or paragraph_element.findall(".//" + qn("w:pict")))


def fix_image_paragraphs(doc, report, space_before_pt=6.0) -> None:
    """Give an inline image its full height and a little air above it.

    pdf2docx can put an image in a paragraph whose line height is exact and
    measured from the text around it. Word then draws the picture inside a
    line box a few points tall, so it overlaps the line above and the top of
    the picture is cut off. Proportional line spacing below 100 percent, which
    is `lineRule="auto"` with a value under 240, crops a tall inline image the
    same way, so both are reset to single spacing here.
    """
    for paragraph in doc.paragraphs:
        element = paragraph._element
        if not has_drawing(element):
            continue
        ppr = element.get_or_add_pPr()
        spacing = ppr.find(qn("w:spacing"))
        if spacing is None:
            spacing = ppr.makeelement(qn("w:spacing"), {})
            ppr.append(spacing)
        rule = spacing.get(qn("w:lineRule"))
        try:
            line = int(spacing.get(qn("w:line")) or 240)
        except ValueError:
            line = 240
        if rule == "exact" or (rule in ("auto", None) and line < 240):
            spacing.set(qn("w:lineRule"), "auto")
            spacing.set(qn("w:line"), "240")
            report["image_paragraphs_fixed"] += 1
        try:
            before = int(spacing.get(qn("w:before")) or 0)
        except ValueError:
            before = 0
        wanted = int(space_before_pt * 20)
        if before < wanted:
            spacing.set(qn("w:before"), str(wanted))
        # An inline picture cannot be wrapped around, so keep it on its own
        # line rather than sharing one with text.
        ppr_keep = ppr.find(qn("w:keepNext"))
        if ppr_keep is None:
            ppr.append(ppr.makeelement(qn("w:keepNext"), {}))


# ---------------------------------------------------------------------------
# 10 November. Publisher tags that were never meant to be read
# ---------------------------------------------------------------------------

# Springer writes these as invisible text beside each equation so that its own
# tooling can find them. pdf_prepare.py takes them out of the PDF before
# pdf2docx runs; this is the safety net for a file that skipped that step, or
# for another publisher's spelling of the same idea.
# Widened 10 November: Springer also writes the inline form as
# "inline-eq-IEq5", which the old pattern missed, so 13 of them reached the
# Word file as text. Kept in step with pdf_prepare.EQ_TOKEN.
def _spaced(word: str) -> str:
    """A word that pdf2docx may have broken with stray spaces.

    Lifting text out of a PDF splits words wherever the glyph spacing looks
    like a gap, so "inline-eq-IEq10" arrives as "inl ine- eq-IEq10". Allowing
    optional space between every letter catches those without loosening the
    match itself: the letters must still be in order, and still be followed by
    the rest of the tag.
    """
    return r"\s*".join(re.escape(ch) for ch in word)


EQ_TOKEN = re.compile(
    # inline-eq-IEq5, display_eq_Equ3, "inl ine- eq- IEq10"
    r"(?i:\b(?:" + _spaced("inline") + r"|" + _spaced("display") + r")"
    r"\s*[-_]\s*e\s*q\s*[-_]?\s*[A-Za-z]*\s*\d*)"
    # Springer's bare inline marker: a capital I, then Eq, then the number.
    # "IEq" is not a word, so this is safe anywhere in a sentence, which
    # matters: 22 of them sat mid-sentence in the 10 November file.
    r"|\bIEqu?\s?\d+\b"
    # A bare "Eq12" or "Equ5" is only a tag when it ends the line. Mid
    # sentence it is far more likely to be the author's own reference, and
    # this document writes those as "Eq. 3".
    r"|\b(?i:Equ?)\d+\b(?=\s*$)",
)

# The tag plus the punctuation that belongs to it. A list of three inline
# symbols reaches Word as "uses TAG, TAG, TAG coefficient values". Deleting
# only the three tags leaves "uses, ,  coefficient values", and tidying that
# afterwards cannot tell which comma was the publisher's and which was ours.
# So each tag is matched together with its own trailing separator and spacing:
# dropping the tag drops them with it, and putting maths back in its place
# keeps them.
EQ_TOKEN_RUN = re.compile(
    r"(?P<tag>" + EQ_TOKEN.pattern + r")(?P<sep>\s*[,;])?(?P<gap>[ \t]*)",
    re.IGNORECASE,
)

# What a deleted tag can still leave behind: a space before punctuation, an
# empty bracket pair, a doubled space.
ORPHAN_SEPARATORS = (
    (re.compile(r"\s*,(\s*,)+"), ","),
    (re.compile(r"\(\s*\)"), ""),
    (re.compile(r"\s+([,.;:)])"), r"\1"),
    (re.compile(r"\(\s+"), "("),
    (re.compile(r"[ \t]{2,}"), " "),
    (re.compile(r"(?:^|(?<=\s))[,;]\s+"), ""),
)


def _sub_across_runs(paragraph, pattern, replacement) -> int:
    """Substitute over a paragraph's whole text without flattening its runs.

    A tag is often split between runs, "inline-eq-" in one and "IEq5" in the
    next, so matching run by run leaves half of it behind. This matches on the
    joined text, maps each match back to the runs it covers, and edits only
    those, so bold and italic elsewhere in the paragraph survive.
    """
    runs = paragraph.runs
    if not runs:
        return 0
    spans, offset = [], 0
    for run in runs:
        spans.append((run, offset, offset + len(run.text)))
        offset += len(run.text)
    whole = "".join(run.text for run in runs)

    matches = list(pattern.finditer(whole))
    if not matches:
        return 0

    # Work out every replacement first, reading left to right. The edits
    # themselves have to run right to left so the offsets ahead of each one
    # stay valid, but the caller hands out maths in document order, so asking
    # for it in reverse gave the first tag the last equation.
    texts = [
        replacement(match) if callable(replacement) else match.expand(replacement)
        for match in matches
    ]

    for match, text in zip(reversed(matches), reversed(texts)):
        first = True
        for run, start, end in spans:
            if end <= match.start() or start >= match.end():
                continue
            cut_from = max(0, match.start() - start)
            cut_to = min(len(run.text), match.end() - start)
            run.text = run.text[:cut_from] + (text if first else "") + run.text[cut_to:]
            first = False
        # Offsets shift once a run is edited, so they are rebuilt each time.
        spans, offset = [], 0
        for run in runs:
            spans.append((run, offset, offset + len(run.text)))
            offset += len(run.text)
    return len(matches)


def _tidy_after_removal(paragraph) -> None:
    for pattern, replacement in ORPHAN_SEPARATORS:
        _sub_across_runs(paragraph, pattern, replacement)


def iter_paragraphs(doc):
    """Every paragraph in the body, including the ones inside tables.

    pdf2docx puts a lot of a laid out page into tables, so a walk over
    `doc.paragraphs` alone misses them. Reading the XML directly keeps nested
    tables in scope as well.
    """
    from docx.text.paragraph import Paragraph

    body = doc.element.body
    for element in body.iter(qn("w:p")):
        yield Paragraph(element, doc)


def _only_paragraph_in_cell(element) -> bool:
    """True when removing this paragraph would leave a table cell with none.

    Word rejects an empty cell, so the last paragraph in one always stays.
    """
    parent = element.getparent()
    if parent is None or parent.tag != qn("w:tc"):
        return False
    return len(parent.findall(qn("w:p"))) <= 1


def drop_eq_tokens(doc, report, replacements=None) -> None:
    """Replace publisher equation tags with the maths they stood for.

    `replacements` is the ordered list `pdf_prepare` collected, one entry per
    tag it could not safely cover in the PDF because the tag is drawn on top
    of the glyphs it names. Covering it there would have taken the beta with
    it, so the tag travels into the Word file and is swapped here instead.

    With nothing to put back, the tag is deleted and the punctuation it
    leaves behind is tidied, so "uses , ,  coefficient values" does not
    happen.
    """
    queue = list(replacements or [])
    position = 0

    def next_replacement(tag):
        nonlocal position
        while position < len(queue):
            entry = queue[position]
            position += 1
            if entry.get("tag") and entry["tag"].strip().lower() != tag.strip().lower():
                # Out of step with the document: keep looking rather than
                # putting the wrong maths in.
                continue
            return str(entry.get("text") or "")
        return ""

    for paragraph in list(iter_paragraphs(doc)):
        if not paragraph.runs:
            continue
        whole = "".join(run.text for run in paragraph.runs)
        if not EQ_TOKEN.search(whole):
            continue
        def swap(match):
            text = next_replacement(match.group("tag"))
            if text:
                report["eq_tokens_replaced"] += 1
                # Keep the separator, so "β1, β2, β3" reads as it did.
                return text + (match.group("sep") or "") + (match.group("gap") or "")
            report["eq_tokens_removed"] += 1
            # Nothing to put back, so the tag goes with the separator that
            # belonged to it. "uses A, B, C coefficient" becomes "uses
            # coefficient" rather than "uses, , coefficient".
            return ""

        _sub_across_runs(paragraph, EQ_TOKEN_RUN, swap)
        # Then tidy what the removal left behind, across runs as well.
        _tidy_after_removal(paragraph)
        # A paragraph that held nothing but tags is now empty furniture.
        if (
            not paragraph.text.strip()
            and not has_drawing(paragraph._element)
            and not _only_paragraph_in_cell(paragraph._element)
        ):
            parent = paragraph._element.getparent()
            if parent is not None:
                parent.remove(paragraph._element)


# ---------------------------------------------------------------------------
# Fault 5. Logo and marker runs
# ---------------------------------------------------------------------------

NO_WORDS = re.compile(r"^[\s\d\W]{0,8}$")


def font_usage(doc) -> Counter:
    used: Counter = Counter()
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            used[split_ps_name(run.font.name or "")[0].lower()] += len(run.text)
    return used


def drop_marker_runs(doc, report, max_chars=40, max_share=0.01) -> None:
    """Remove a paragraph that is only page furniture drawn in a rare font.

    A page number set in the body font is never touched, because the body
    font carries most of the document.
    """
    used = font_usage(doc)
    total = max(1, sum(used.values()))
    rare = {name for name, count in used.items()
            if name and count <= max_chars and count / total <= max_share}
    if not rare:
        return
    for paragraph in list(doc.paragraphs):
        text = paragraph.text.strip()
        runs = [run for run in paragraph.runs if run.text.strip()]
        if not text or not runs or has_drawing(paragraph._element):
            continue
        families = {split_ps_name(run.font.name or "")[0].lower() for run in runs}
        families.discard("")
        if not families or not families <= rare or not NO_WORDS.match(text):
            continue
        report["markers_removed"].append({"text": text, "fonts": sorted(families)})
        paragraph._element.getparent().remove(paragraph._element)


# ---------------------------------------------------------------------------
# Faults 1 and 2, review points b and c. Running heads and feet
# ---------------------------------------------------------------------------

def section_paragraphs(doc):
    """Body paragraphs grouped by the section they belong to, in page order."""
    from docx.text.paragraph import Paragraph

    groups, current = [], []
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            current.append(Paragraph(child, doc))
            ppr = child.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                groups.append(current)
                current = []
        elif tag == "sectPr":
            groups.append(current)
            current = []
    if current:
        groups.append(current)
    return groups


MIN_RUNNING_SHARE = 0.30


def _needed(slots: int, share: float = MIN_RUNNING_SHARE) -> int:
    """How many pages a line must appear on before it counts as a running one."""
    return max(2, -(-int(slots) * int(share * 100) // 100))


def _classify_running(lines, share: float = MIN_RUNNING_SHARE):
    """Decide whether a repeated line is the same on every page or alternates.

    Returns (kind, odd_text, even_text) where kind is "none", "same" or
    "alternating". Page 1 counts as odd, which is what Word means by the
    default header.

    Judged by the share of pages a line covers, not by an all or nothing test.
    The 10 November test file showed why both halves of that matter:

      Too strict. A journal article's first page carries the banner and the
      title block, not a running head, so the odd pages hold two different
      first lines. The old test demanded exactly one, gave up, and left the
      head in the body on 17 of 35 pages.

      Too loose. After the publisher's logo lines were blanked, only two
      pages had a bottom line left, both the title head. Two out of 35 was
      enough for the old "same" branch, so the head was written into the
      footer of every page while the header stayed empty.

    A line now has to cover at least `share` of the pages it could appear on,
    and never fewer than two.
    """
    present = [text for text in lines if text]
    total = len(lines)
    if total < 2 or len(present) < 2:
        return "none", None, None

    odd_lines = [lines[index] for index in range(0, total, 2)]
    even_lines = [lines[index] for index in range(1, total, 2)]
    odd_counts = Counter(text for text in odd_lines if text)
    even_counts = Counter(text for text in even_lines if text)

    odd_top, odd_hits = odd_counts.most_common(1)[0] if odd_counts else (None, 0)
    even_top, even_hits = even_counts.most_common(1)[0] if even_counts else (None, 0)

    # Left and right pages carrying different heads, which is what a journal
    # article does. Checked first, so a file that alternates is never flattened
    # into one head for every page.
    if (odd_top and even_top and odd_top != even_top
            and odd_hits >= _needed(len(odd_lines), share)
            and even_hits >= _needed(len(even_lines), share)):
        return "alternating", odd_top, even_top

    counts = Counter(present)
    most_common, hits = counts.most_common(1)[0]
    if hits >= _needed(total, share):
        return "same", most_common, None
    return "none", None, None


# A caption is the thing most likely to sit at the bottom of a page and look
# repeated, especially in a file made by merging the same pages twice.
CAPTION_START = re.compile(
    r"^\s*(fig(?:ure)?|tab(?:le)?|chart|scheme|exhibit|plate|box|appendix|source|notes?)"
    r"\b[\s:.]*\d*[\s:.]",
    re.IGNORECASE,
)


def _is_caption(text: str) -> bool:
    """True when a line opens with a caption or source label."""
    return bool(CAPTION_START.match(str(text or "")))


def _is_sentencey(text: str) -> bool:
    """True when a line reads like running prose rather than page furniture.

    Six words or more, at least half of them four letters or longer. A journal
    name, an author line and a page number all fall short of that; a sentence
    does not.
    """
    words = str(text or "").split()
    if len(words) < 6:
        return False
    long_words = sum(1 for word in words if len(re.sub(r"[^A-Za-z]", "", word)) >= 4)
    return long_words / len(words) >= 0.5


def _set_part_text(part, text, alignment):
    part.is_linked_to_previous = False
    paragraph = part.paragraphs[0]
    paragraph.text = text
    if alignment is not None:
        try:
            paragraph.alignment = alignment
        except ValueError:
            pass


def lift_running_text(doc, report, max_len=140):
    """Move repeated top and bottom lines into real headers and footers.

    When the two running heads alternate between left and right pages, as a
    journal article does, the document is switched to different odd and even
    page headers so both are kept and each lands on the right page. Later
    sections link back to the first, so one pair of headers covers the file.

    Returns the extra top and bottom space each section needs.
    """
    def edges():
        """First and last paragraph with words on it, per section, now."""
        tops, bottoms = [], []
        for group in section_paragraphs(doc):
            filled = [p for p in group if p.text.strip()]
            tops.append(filled[0] if filled else None)
            bottoms.append(filled[-1] if filled else None)
        return tops, bottoms

    def text_of(paragraph):
        if paragraph is None:
            return ""
        text = paragraph.text.strip()
        return text if 0 < len(text) <= max_len else ""

    tops, bottoms = edges()
    if not tops:
        return {}

    # A repeated line with no words in it is page furniture, not a running
    # head. The "1 3" that the Springer logo leaves behind is removed here.
    #
    # The edges are deliberately not read again afterwards. Taking the logo
    # away uncovers whatever sat above it, which on this article is the figure
    # caption, and a caption that repeats is not a running foot. The sparse
    # bottom list that the logo leaves behind is handled by the support
    # threshold in `_classify_running`, which is what the 10 November footer
    # bug actually needed.
    for candidates in (tops, bottoms):
        lines = [text_of(p) for p in candidates]
        repeated = {text for text, n in Counter(t for t in lines if t).items() if n >= 2}
        for index, text in enumerate(lines):
            if text and text in repeated and not re.search(r"[A-Za-z]{2,}", text):
                paragraph = candidates[index]
                if paragraph is None or paragraph._element.getparent() is None:
                    continue
                report["markers_removed"].append({"text": text, "fonts": ["repeated marker"]})
                paragraph._element.getparent().remove(paragraph._element)
                candidates[index] = None

    top_lines = [text_of(p) for p in tops]
    bottom_lines = [text_of(p) for p in bottoms]

    head_kind, head_odd, head_even = _classify_running(top_lines)
    foot_kind, foot_odd, foot_even = _classify_running(bottom_lines)

    # A line is a head or a foot, never both. The top wins, because that is
    # where a running head belongs and where it was read from.
    head_texts = {text for text in (head_odd, head_even) if text}
    if foot_kind != "none" and head_texts:
        foot_texts = {text for text in (foot_odd, foot_even) if text}
        if foot_texts & head_texts:
            report["footer_rejected"] = sorted(foot_texts & head_texts)
            foot_kind, foot_odd, foot_even = "none", None, None

    # A running foot is page furniture: a number, a journal line, a notice. A
    # sentence at the bottom of a page is a caption or the end of a paragraph,
    # and on a short document two of them can look alike. Prose needs three
    # pages of evidence before it is believed.
    if foot_kind != "none":
        candidates_text = [text for text in (foot_odd, foot_even) if text]
        counts = Counter(text for text in bottom_lines if text)
        if any(_is_caption(text) or (_is_sentencey(text) and counts.get(text, 0) < 3)
               for text in candidates_text):
            report["footer_rejected"] = sorted(set(report["footer_rejected"]) | set(candidates_text))
            foot_kind, foot_odd, foot_even = "none", None, None

    if head_kind == "none" and foot_kind == "none":
        return {}

    alternating = "alternating" in (head_kind, foot_kind)
    if alternating:
        doc.settings.odd_and_even_pages_header_footer = True

    sections = doc.sections
    first = sections[0]
    extra: dict[int, dict[str, float]] = {}

    def line_height(paragraph, fallback=10.0):
        sizes = [r.font.size.pt for r in paragraph.runs if r.font.size is not None]
        return (max(sizes) if sizes else fallback) * 1.6

    # Write the header and footer parts once, on the first section.
    for kind, odd_text, even_text, part_name in (
        (head_kind, head_odd, head_even, "header"),
        (foot_kind, foot_odd, foot_even, "footer"),
    ):
        if kind == "none":
            continue
        source = tops if part_name == "header" else bottoms
        sample = next((p for p in source if p is not None), None)
        alignment = sample.alignment if sample is not None else None

        default_part = getattr(first, part_name)
        _set_part_text(default_part, odd_text, alignment)
        if kind == "alternating":
            even_part = getattr(first, "even_page_" + part_name)
            _set_part_text(even_part, even_text, alignment)
            report[part_name + "s"] = {"odd": odd_text, "even": even_text}
        else:
            report[part_name + "s"] = {"every page": odd_text}

    for section in sections[1:]:
        for part_name in ("header", "footer"):
            try:
                getattr(section, part_name).is_linked_to_previous = True
                if alternating:
                    getattr(section, "even_page_" + part_name).is_linked_to_previous = True
            except (AttributeError, ValueError):
                pass

    # Now take the lines out of the body, but only from the pages that really
    # carry them. A first page with a banner instead of a running head keeps
    # its banner.
    foot_texts = {text for text in (foot_odd, foot_even) if text}
    for index, (top, bottom) in enumerate(zip(tops, bottoms)):
        if (head_kind != "none" and top is not None
                and top_lines[index] in head_texts
                and top._element.getparent() is not None):
            extra.setdefault(index, {})["header"] = line_height(top)
            top._element.getparent().remove(top._element)
            report["lines_lifted"] += 1
        if (foot_kind != "none" and bottom is not None and bottom is not top
                and bottom_lines[index] in foot_texts
                and bottom._element.getparent() is not None):
            extra.setdefault(index, {})["footer"] = line_height(bottom)
            bottom._element.getparent().remove(bottom._element)
            report["lines_lifted"] += 1

    return extra


# ---------------------------------------------------------------------------
# Fault 4. Line end hyphens
# ---------------------------------------------------------------------------

HYPHEN_WORD = re.compile(r"\b([A-Za-z]{2,})-([a-z]{2,})\b")
WORD = re.compile(r"[A-Za-z]{2,}")


def build_vocabulary(doc) -> Counter:
    """Count words that stand on their own.

    Hyphenated pairs come out first, so a broken half such as "ment" can
    never be evidence that "ment" is a word.
    """
    words: Counter = Counter()
    for paragraph in doc.paragraphs:
        for word in WORD.findall(HYPHEN_WORD.sub(" ", paragraph.text)):
            words[word.lower()] += 1
    return words


def rejoin_hyphens(doc, report) -> None:
    vocab = build_vocabulary(doc)
    seen_joined, seen_kept = set(), set()

    def decide(match):
        left, right = match.group(1), match.group(2)
        joined = left + right
        whole = match.group(0)
        # 1. The joined word already appears elsewhere without a hyphen.
        if vocab.get(joined.lower(), 0) > 0:
            if whole not in seen_joined:
                seen_joined.add(whole)
                report["hyphens_joined"].append(whole + " to " + joined)
            return joined
        # 2. Neither half stands on its own anywhere in the document, and the
        #    left half is long enough and not an abbreviation, so the hyphen
        #    split one word. "manu-facturing" joins, "energy-saving" does not.
        if (len(left) >= 4 and not left.isupper()
                and vocab.get(left.lower(), 0) == 0
                and vocab.get(right.lower(), 0) == 0):
            if whole not in seen_joined:
                seen_joined.add(whole)
                report["hyphens_joined"].append(whole + " to " + joined)
            return joined
        if whole not in seen_kept:
            seen_kept.add(whole)
            report["hyphens_kept"].append(whole)
        return whole

    for paragraph in doc.paragraphs:
        runs = paragraph.runs
        if not runs:
            continue
        joined_text = "".join(run.text for run in runs)
        if not HYPHEN_WORD.search(joined_text):
            continue
        fixed = HYPHEN_WORD.sub(decide, joined_text)
        if fixed == joined_text:
            continue
        formats = {(run.font.name, run.bold, run.italic) for run in runs}
        if len(formats) == 1:
            runs[0].text = fixed
            for run in runs[1:]:
                run.text = ""
            continue
        # Mixed formatting: work on each run and on each pair of neighbours,
        # so a hyphen that sits on a run boundary is still caught.
        for index, run in enumerate(runs):
            run.text = HYPHEN_WORD.sub(decide, run.text)
        for index in range(len(runs) - 1):
            pair = runs[index].text + runs[index + 1].text
            new_pair = HYPHEN_WORD.sub(decide, pair)
            if new_pair != pair:
                cut = len(new_pair) - len(runs[index + 1].text)
                runs[index].text = new_pair[:max(0, cut)]
                runs[index + 1].text = new_pair[max(0, cut):]


# ---------------------------------------------------------------------------
# Fault 6. Hard breaks and broken paragraphs
# ---------------------------------------------------------------------------

SENTENCE_END = re.compile(r"[.!?:;…][\"'”’)\]]?\s*$")


def breaks_to_spaces(doc, report) -> None:
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for element in run._element.findall(qn("w:br")):
                if element.get(qn("w:type")) in ("page", "column"):
                    continue
                run._element.remove(element)
                report["breaks_removed"] += 1
                if run.text and not run.text.endswith(" "):
                    run.text += " "


def merge_broken_paragraphs(doc, report, min_length=40) -> None:
    """Join a paragraph that was cut mid sentence onto the one before it."""
    from docx.text.paragraph import Paragraph

    body = doc.element.body
    previous = None
    for element in [c for c in body.iterchildren() if c.tag == qn("w:p")]:
        paragraph = Paragraph(element, doc)
        text = paragraph.text
        if not text.strip():
            continue
        ppr = element.find(qn("w:pPr"))
        ends_section = ppr is not None and ppr.find(qn("w:sectPr")) is not None
        if previous is None or has_drawing(element) or has_drawing(previous._element):
            previous = paragraph
            continue
        previous_text = previous.text
        joinable = (previous_text.strip()
                    and len(previous_text) >= min_length
                    and not SENTENCE_END.search(previous_text)
                    and text[:1].islower()
                    and not ends_section)
        if not joinable:
            previous = paragraph
            continue
        separator = "" if previous_text.endswith(("-", "‐")) else " "
        if previous.runs and separator:
            previous.runs[-1].text += separator
        for run in paragraph.runs:
            previous._element.append(run._element)
        element.getparent().remove(element)
        report["paragraphs_merged"] += 1


# ---------------------------------------------------------------------------
# Fault 1. One flowing document, margins, blank space
# ---------------------------------------------------------------------------

def merge_sections(doc, report) -> None:
    """Keep only the final sectPr, so the text flows instead of breaking.

    The running heads were written to the first section, so its header and
    footer references move to the section that survives. Without this the
    whole document comes back with no header at all.
    """
    body = doc.element.body
    inline = []
    for paragraph in list(body.iterchildren(qn("w:p"))):
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            continue
        sect_pr = ppr.find(qn("w:sectPr"))
        if sect_pr is not None:
            inline.append((paragraph, ppr, sect_pr))

    final = body.find(qn("w:sectPr"))
    if inline and final is not None:
        first = inline[0][2]
        references = [
            child
            for child in first
            if child.tag in (qn("w:headerReference"), qn("w:footerReference"))
        ]
        if references:
            for child in list(final):
                if child.tag in (qn("w:headerReference"), qn("w:footerReference")):
                    final.remove(child)
            # Both references must come before everything else in a sectPr.
            for child in reversed(references):
                first.remove(child)
                final.insert(0, child)

    for paragraph, ppr, sect_pr in inline:
        ppr.remove(sect_pr)
        report["sections_merged"] += 1
        if not len(ppr):
            paragraph.remove(ppr)


def drop_empty_paragraphs(doc, report, keep=0) -> None:
    """Remove blank paragraphs that pdf2docx left behind.

    These are not spacing the author asked for. In the 10 November file there
    were 71 of them, each carrying an exact line height of 11 pt and 22 pt of
    space after it: about 33 pt of nothing apiece, which is close to three
    pages across the document. Keeping one of each run, as this used to, kept
    all 71.

    A paragraph holding a picture is never blank, and the one that carries a
    section break is always left alone.
    """
    from docx.text.paragraph import Paragraph

    for element in list(doc.element.body.iterchildren(qn("w:p"))):
        paragraph = Paragraph(element, doc)
        ppr = element.find(qn("w:pPr"))
        ends_section = ppr is not None and ppr.find(qn("w:sectPr")) is not None
        if paragraph.text.strip() or has_drawing(element) or ends_section:
            continue
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)
            report["blank_paragraphs_removed"] += 1


def _max_font_pt(paragraph, fallback=0.0) -> float:
    sizes = [run.font.size.pt for run in paragraph.runs if run.font.size is not None]
    return max(sizes) if sizes else fallback


def normalise_spacing(doc, report, max_gap_pt=8.0) -> None:
    """Even out line spacing and cap the gaps between paragraphs.

    pdf2docx measures a line height per paragraph from the PDF, so a document
    comes back with dozens of slightly different multipliers, and after an
    equation the measurements drift far enough that the text looks almost
    double spaced. In flowing text none of that carries meaning, so every
    paragraph gets single spacing and the space before and after is capped.

    A paragraph holding a picture is left exactly as it is: proportional
    spacing below 100 percent crops a tall inline image.

    Keep layout wants this too, with a looser cap on the gaps. The 35 page
    paper came back as 44 pages. About three of those were blank paragraphs,
    and the rest was this: `max_line_spacing_ratio` lets pdf2docx write up to
    125 percent, measured from the line pitch rather than asked for by the
    author, and 15 percent of extra height over 35 pages is five more pages.
    Single spacing is what the paper itself uses, so this moves the result
    towards the original rather than away from it.
    """
    from docx.text.paragraph import Paragraph

    for element in doc.element.body.iter(qn("w:p")):
        if has_drawing(element):
            continue
        paragraph = Paragraph(element, doc)
        if not paragraph.text.strip():
            continue
        ppr = element.find(qn("w:pPr"))
        if ppr is None:
            continue
        spacing = ppr.find(qn("w:spacing"))
        if spacing is None:
            continue
        changed = False
        rule = spacing.get(qn("w:lineRule"))
        line = spacing.get(qn("w:line"))
        # Exact spacing smaller than the text would clip it, so that goes too.
        too_tight = (rule == "exact" and line
                     and int(line) < _max_font_pt(paragraph, 9.0) * 20 * 1.15)
        if rule != "auto" or (line and line != "240") or too_tight:
            spacing.set(qn("w:lineRule"), "auto")
            spacing.set(qn("w:line"), "240")
            changed = True
        for side in ("w:before", "w:after"):
            value = spacing.get(qn(side))
            if value and int(value) > max_gap_pt * 20:
                spacing.set(qn(side), str(int(max_gap_pt * 20)))
                changed = True
        if changed:
            report["spacing_normalised"] += 1


def drop_page_breaks(doc, report) -> None:
    """Take out page breaks, which mean nothing once the text flows."""
    for element in doc.element.body.iter(qn("w:p")):
        for br in element.findall(".//" + qn("w:br")):
            if br.get(qn("w:type")) == "page":
                parent = br.getparent()
                if parent is not None:
                    parent.remove(br)
                    report["page_breaks_removed"] += 1


def set_page_setup(doc, extra=None, floor=18.0) -> None:
    """Keep the PDF page size and give the lifted lines room above and below."""
    extra = extra or {}
    for index, section in enumerate(doc.sections):
        room = extra.get(index, {})
        top = section.top_margin.pt if section.top_margin is not None else floor
        bottom = section.bottom_margin.pt if section.bottom_margin is not None else floor
        header_gap = max(floor / 2, min(top, 36.0))
        footer_gap = max(floor / 2, min(bottom, 36.0))
        if "header" in room:
            top = header_gap + room["header"]
        if "footer" in room:
            bottom = footer_gap + room["footer"]
        section.top_margin = Pt(max(top, floor))
        section.bottom_margin = Pt(max(bottom, floor))
        section.header_distance = Pt(max(6.0, min(header_gap, section.top_margin.pt - 6)))
        section.footer_distance = Pt(max(6.0, min(footer_gap, section.bottom_margin.pt - 6)))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def new_report() -> dict:
    return {
        "mode": "",
        "fonts": {},
        "columns_fixed": 0,
        "column_counts": [],
        "image_paragraphs_fixed": 0,
        "markers_removed": [],
        "headers": {},
        "footers": {},
        "footer_rejected": [],
        "lines_lifted": 0,
        "eq_tokens_removed": 0,
        "eq_tokens_replaced": 0,
        "spacing_normalised": 0,
        "page_breaks_removed": 0,
        "hyphens_joined": [],
        "hyphens_kept": [],
        "breaks_removed": 0,
        "paragraphs_merged": 0,
        "sections_merged": 0,
        "blank_paragraphs_removed": 0,
    }


def repair(in_path: str, out_path: str, mode: str = "flow",
           inline_math=None) -> dict:
    """Repair a pdf2docx DOCX. Returns the report.

    `inline_math` is the ordered list `pdf_prepare` collected for tags it
    could not cover in the PDF without taking the maths glyphs with them.
    """
    document = Document(in_path)
    report = new_report()
    report["mode"] = mode

    roots = [document.element.body, document.styles.element]
    for section in document.sections:
        for part_name in ("header", "footer", "even_page_header", "even_page_footer"):
            part = getattr(section, part_name, None)
            if part is not None:
                roots.append(part._element)
    fix_fonts(roots, report)

    fix_columns(document, report)
    fix_image_paragraphs(document, report)
    drop_eq_tokens(document, report, inline_math)
    drop_marker_runs(document, report)
    extra = lift_running_text(document, report)
    rejoin_hyphens(document, report)
    breaks_to_spaces(document, report)

    # Blank paragraphs are pdf2docx leftovers in both modes. They carry no
    # words and no picture, so all they add is height, and in Keep layout
    # that height is what pushes a page onto the next one.
    drop_empty_paragraphs(document, report)

    if mode == "flow":
        merge_broken_paragraphs(document, report)
        drop_page_breaks(document, report)
        normalise_spacing(document, report)
        merge_sections(document, report)
    else:
        # Keep layout gets the same line spacing, with a looser cap on the
        # gaps between paragraphs so the shape of the page survives. Without
        # this the blind tighten has to find 26 percent, and 3 percent is all
        # it can safely take.
        normalise_spacing(document, report, max_gap_pt=12.0)

    set_page_setup(document, extra)
    document.save(out_path)
    return report


if __name__ == "__main__":
    import json
    import sys

    chosen = sys.argv[3] if len(sys.argv) > 3 else "flow"
    print(json.dumps(repair(sys.argv[1], sys.argv[2], chosen), indent=2))
