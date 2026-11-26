import os
import shutil
import subprocess

import pdfplumber
import re
from collections import Counter


class PolicyPDFReader:

    PAGE_NUM_RE = re.compile(r"^Page\s+\d+\s+of\s+\d+\s*$", re.IGNORECASE)
    HEADING_RE = re.compile(r"^\s*(\d+(?:\.\d+)*)\.?\s+([^\d].{2,200})$")
    BULLET_RE = re.compile(r"^\s*(?:[•\-\*]|[A-Za-z]\)|\d+\))\s+")

    OCR_THRESHOLD_CHARS_PER_PAGE = 30

    @classmethod
    def read(cls, pdf_path):

        text_per_page = cls._extract_pages(pdf_path)

        avg_chars = (
            sum(len(t) for t in text_per_page) / max(1, len(text_per_page))
        )

        if avg_chars < cls.OCR_THRESHOLD_CHARS_PER_PAGE:

            ocr_path = cls._try_ocr(pdf_path)

            if ocr_path:
                print(f"    ↻ used OCR for {os.path.basename(pdf_path)}")
                text_per_page = cls._extract_pages(ocr_path)

        raw_lines = []

        for text in text_per_page:

            for line in text.split("\n"):

                line = line.rstrip()

                if cls.PAGE_NUM_RE.match(line):
                    continue

                raw_lines.append(line)

        return cls._to_paragraphs(raw_lines)

    @classmethod
    def _extract_pages(cls, pdf_path):

        with pdfplumber.open(pdf_path) as pdf:

            return [cls._page_text(p) for p in pdf.pages]

    # Two column detection. pdfplumber's extract_text() reads straight across
    # the page, so on a two column page every line is "left line + right
    # line" and sentences from different columns are woven together
    # (confirmed on China's 182 page MOFCOM guide, 131 pages two column: a
    # contact phone number landed in the middle of a sentence about Shanghai
    # headquarters). Such pages are read one column at a time, left column
    # first, which is also how a person reads them.
    #
    # Deliberately conservative, because a wrong split is worse than the
    # status quo: a page is only treated as two column when
    #   * a clear empty vertical gutter sits near the centre,
    #   * words crossing it are at most a few and all in a heading band at
    #     the top or a footer band at the bottom (a crossing word mid page
    #     means a table, figure or irregular layout), and
    #   * BOTH sides hold many lines of prose length text (a two column
    #     table of short cells would otherwise be split, separating each
    #     label from its value), and
    #   * ruled tables cover at most a quarter of the page. Known residual
    #     risk: a table with no ruling lines and long text cells on both
    #     sides still passes (Morocco's regional totals page), and is then
    #     read column by column.
    # Anything else is read exactly as before. A full width heading or
    # footer band is read on its own, before / after the two columns.
    GUTTER_BAND = (0.35, 0.65)
    GUTTER_MIN_WIDTH = 6.0
    GUTTER_MAX_CROSSING_WORDS = 4
    HEADER_ZONE = 0.20
    FOOTER_ZONE = 0.90
    COLUMN_MIN_WORDS_ON_PAGE = 80
    COLUMN_MIN_LINES = 10
    COLUMN_MIN_MEDIAN_WORDS_PER_LINE = 4
    COLUMN_MAX_WORD_DIFF = 2
    BAND_MARGIN = 0.5
    COLUMN_MAX_TABLE_SHARE = 0.25

    @classmethod
    def _page_text(cls, page):

        layout = cls._find_layout(page)

        if layout is None:
            return page.extract_text() or ""

        gutter, header_end, footer_start = layout

        def text_of(x0, top, x1, bottom):
            return page.crop((x0, top, x1, bottom)).extract_text() or ""

        parts = []

        if header_end is not None:
            parts.append(text_of(0, 0, page.width, header_end))

        body_top = header_end if header_end is not None else 0
        body_bottom = footer_start if footer_start is not None else page.height

        parts.append(text_of(0, body_top, gutter, body_bottom))
        parts.append(text_of(gutter, body_top, page.width, body_bottom))

        if footer_start is not None:
            parts.append(text_of(0, footer_start, page.width, page.height))

        column_text = "\n".join(p for p in parts if p)

        # Safety net: a column read must contain exactly the same words as
        # the plain read, only in a different order. A crop that duplicated
        # or split a word (a title straddling the gutter, a gutter narrower
        # than pdfplumber's word gap) changes the word counts, and the page
        # then falls back to the plain read rather than risk corrupt text.
        plain_text = page.extract_text() or ""

        before = Counter(plain_text.split())
        after = Counter(column_text.split())

        if sum(((before - after) + (after - before)).values()) > cls.COLUMN_MAX_WORD_DIFF:
            return plain_text

        return column_text

    @staticmethod
    def _line_extent(words, seed):
        """Vertical extent (top, bottom) of the text line containing seed:
        grown until no word whose centre falls inside sticks out of it."""

        top, bottom = seed["top"], seed["bottom"]

        changed = True

        while changed:

            changed = False

            for w in words:

                centre = (w["top"] + w["bottom"]) / 2

                if top <= centre <= bottom and (w["top"] < top or w["bottom"] > bottom):
                    top = min(top, w["top"])
                    bottom = max(bottom, w["bottom"])
                    changed = True

        return top, bottom

    @classmethod
    def _find_layout(cls, page):
        """(gutter_x, header_end_y or None, footer_start_y or None) if the
        page is clearly two column, else None."""

        try:
            words = page.extract_words()
        except Exception:
            return None

        if len(words) < cls.COLUMN_MIN_WORDS_ON_PAGE:
            return None

        lo = int(page.width * cls.GUTTER_BAND[0])
        hi = int(page.width * cls.GUTTER_BAND[1])

        # Words covering each integer x in the central band; the gutter is
        # the widest run covered by at most a few words (the crossing
        # heading / footer words allowed for below).
        coverage = [0] * (hi - lo + 1)

        for w in words:
            for x in range(max(lo, int(w["x0"])), min(hi, int(w["x1"])) + 1):
                coverage[x - lo] += 1

        best = None
        run_start = None

        for i, c in enumerate(coverage + [cls.GUTTER_MAX_CROSSING_WORDS + 1]):

            if c <= cls.GUTTER_MAX_CROSSING_WORDS:
                if run_start is None:
                    run_start = i
            elif run_start is not None:
                if best is None or i - run_start > best[1] - best[0]:
                    best = (run_start, i)
                run_start = None

        if not best or best[1] - best[0] < cls.GUTTER_MIN_WIDTH:
            return None

        gutter = lo + (best[0] + best[1]) / 2

        crossing = [w for w in words if w["x0"] < gutter < w["x1"]]

        header_end = None
        footer_start = None

        for w in crossing:

            centre = (w["top"] + w["bottom"]) / 2

            # The band must cover the WHOLE text line the crossing word is
            # on (taller words on the same line stick out past it), or the
            # rest of that line is read again as part of the column body.
            line_top, line_bottom = cls._line_extent(words, w)

            # 0.5 pt past the line: pdfplumber's crop counts a character
            # that only touches the edge as overlapping, so a boundary
            # exactly on the line edge re-reads the heading in clipped
            # pieces ("Muni" / "icipalities)").
            if centre < page.height * cls.HEADER_ZONE:
                header_end = max(header_end or 0, line_bottom + cls.BAND_MARGIN)
            elif centre > page.height * cls.FOOTER_ZONE:
                footer_start = min(footer_start or page.height, line_top - cls.BAND_MARGIN)
            else:
                return None

        def in_band(w):
            return (
                (header_end is not None and w["top"] < header_end)
                or (footer_start is not None and w["bottom"] > footer_start)
            )

        body = [w for w in words if not in_band(w)]

        for side in (
            [w for w in body if w["x1"] <= gutter],
            [w for w in body if w["x0"] >= gutter],
        ):

            lines = {}

            for w in side:
                lines.setdefault(round(w["top"] / 3), []).append(w)

            counts = sorted(len(v) for v in lines.values())

            if len(counts) < cls.COLUMN_MIN_LINES:
                return None

            if counts[len(counts) // 2] < cls.COLUMN_MIN_MEDIAN_WORDS_PER_LINE:
                return None

        # Last, because it is the slowest check: a page that is largely a
        # ruled table is not prose in columns. Reading it column by column
        # separates each row's label from its value (Romania's payment
        # request document table, confirmed). Page failures to detect
        # tables are treated as "not safe to split".
        try:
            table_area = sum(
                (t.bbox[2] - t.bbox[0]) * (t.bbox[3] - t.bbox[1])
                for t in page.find_tables()
            )
        except Exception:
            return None

        if table_area / (page.width * page.height) > cls.COLUMN_MAX_TABLE_SHARE:
            return None

        return gutter, header_end, footer_start

    @staticmethod
    def _try_ocr(pdf_path):

        if not shutil.which("ocrmypdf"):
            return None

        ocr_path = pdf_path.replace(".pdf", ".ocr.pdf")

        if os.path.exists(ocr_path):
            return ocr_path

        try:

            subprocess.run(
                [
                    "ocrmypdf", "--skip-text", "--optimize", "1",
                    "--quiet", pdf_path, ocr_path,
                ],
                check=True,
                timeout=600,
            )

            return ocr_path

        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):

            return None

    @classmethod
    def _to_paragraphs(cls, lines):

        paragraphs = []

        current_heading = ""

        current_buffer = []

        def flush():

            if current_buffer:

                text = " ".join(b.strip() for b in current_buffer if b.strip())

                if text and len(text) > 15:

                    paragraphs.append({
                        "heading": current_heading,
                        "text": text
                    })

        for line in lines:

            stripped = line.strip()

            if not stripped:

                flush()
                current_buffer = []
                continue

            m = cls.HEADING_RE.match(line)

            if m and len(stripped) < 100 and not cls.BULLET_RE.match(line):

                flush()
                current_buffer = []
                current_heading = f"{m.group(1)} {m.group(2)}".strip()

                continue

            current_buffer.append(stripped)

        flush()

        return paragraphs
