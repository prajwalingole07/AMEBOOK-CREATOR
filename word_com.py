"""
MS Word formatting pass for the CET college cut-off files.

Ported from the JoSAA book's word_com.py: only Word itself can lay pages
out, so every generated .docx is optionally opened (hidden, via the COM
automation model - pywin32) and given the same treatment as the JoSAA
downloads:

  1. fit_college_tables() - each rank table (and the info box) is forced to
     span exactly the usable page width (100% of the text area), centred, so
     it can never run past the margins; then every rank table that spills
     onto a second page is shrunk (font size) and re-checked, up to
     MAX_PASSES times, exactly like the JoSAA one-page audit.
  2. fit_bytes()           - in-memory wrapper used by /api/download-docx.

All entry points degrade gracefully: on Vercel / Linux / machines without
pywin32 or Word, fit_bytes() returns the file untouched (the python-docx
layout already matches the JoSAA numbers: 0.4" margins, 10.85" tables,
centred).
"""

import os
import tempfile

try:
    import win32com.client
except Exception:                     # Vercel, Linux, pywin32 missing
    win32com = None

WD_ACTIVE_END_PAGE_NUMBER = 3      # WdInformation.wdActiveEndPageNumber
WD_DO_NOT_SAVE_CHANGES = 0
WD_SAVE_CHANGES = -1
WD_AUTO_FIT_FIXED = 0              # WdAutoFitBehavior.wdAutoFitFixed
WD_PREFERRED_WIDTH_POINTS = 3      # WdPreferredWidthType.wdPreferredWidthPoints
WD_ROW_ALIGN_CENTER = 1            # WdRowAlignment.wdAlignRowCenter

MIN_FONT_PT = 5.5                  # never shrink a table below this size
MAX_PASSES = 6
SHRINK_STEP = 0.5                  # points taken off per pass

DOC_TABLE_WIDTH_IN = 10.85         # same table width the JoSAA files use


# --------------------------------------------------------------------------- #
#  helpers
# --------------------------------------------------------------------------- #

def is_available():
    """True when this machine can actually run the Word formatting pass."""
    if os.environ.get("VERCEL"):
        return False
    if os.name != "nt" or win32com is None:
        return False
    return True


def _clean(txt):
    return (txt or "").replace("\x07", "").replace("\r", "").replace("\x0b", " ").strip()


def _is_rank_table(table):
    """A CET rank table starts with the 'Branch (Intake)' header cell.
    The info box starts with 'College Code' - that one is handled too but
    never shrunk."""
    try:
        return _clean(table.Cell(1, 1).Range.Text).startswith("Branch")
    except Exception:
        return False


def _is_info_box(table):
    try:
        return _clean(table.Cell(1, 1).Range.Text).startswith("College Code")
    except Exception:
        return False


def _table_title(table):
    """Title paragraph sitting directly above the rank table ('Home
    University Cut-Off' etc.) - used only for progress messages."""
    try:
        prev = table.Range.Previous(1, 1)          # wdParagraph = 1
        return _clean(prev.Text)
    except Exception:
        return ""


def _page_of(rng):
    return int(rng.Information(WD_ACTIVE_END_PAGE_NUMBER))


def _start_page_of(table):
    return _page_of(table.Cell(1, 1).Range)


def _end_page_of(table):
    return _page_of(table.Range)


def _fit_width(table):
    """Make the table exactly DOC_TABLE_WIDTH_IN wide, centred.

    This is the COM half of the JoSAA layout guarantee: Word itself rescales
    the columns to the fixed 10.85" preferred width (plain twip widths, the
    same numbers the JoSAA files carry) and centres the table, so it can
    never extend past the page margins.
    """
    try:
        table.PreferredWidthType = WD_PREFERRED_WIDTH_POINTS
        table.PreferredWidth = DOC_TABLE_WIDTH_IN * 72      # twips/points
        table.AutoFitBehavior(WD_AUTO_FIT_FIXED)
    except Exception:
        pass
    try:
        table.Rows.Alignment = WD_ROW_ALIGN_CENTER
    except Exception:
        pass


def _shrink_table(table):
    """Take SHRINK_STEP points off every paragraph of the table.

    Returns False when the table is already at MIN_FONT_PT.
    """
    sizes = []
    for para in table.Range.Paragraphs:
        try:
            size = para.Range.Font.Size
        except Exception:
            size = None
        if size is None:            # mixed runs inside one paragraph
            size = MIN_FONT_PT + SHRINK_STEP
        sizes.append(size)
    if not sizes or min(sizes) - SHRINK_STEP < MIN_FONT_PT:
        return False

    for para in table.Range.Paragraphs:
        try:
            size = para.Range.Font.Size
            if size is None:
                size = MIN_FONT_PT + SHRINK_STEP
            para.Range.Font.Size = max(MIN_FONT_PT, size - SHRINK_STEP)
        except Exception:
            pass
        try:
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
        except Exception:
            pass
    return True


def open_word():
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


# --------------------------------------------------------------------------- #
#  one file: fit width + one page per rank table
# --------------------------------------------------------------------------- #

def fit_college_tables(path, progress_cb=None):
    """Format one college .docx in place, JoSAA style:

    * every table is fitted to the usable width and centred (never extends
      past the margins);
    * every rank table that spills onto a second page is shrunk until it
      fits on one landscape A4 page, up to MAX_PASSES times.
    """
    if not is_available():
        return {"checked": 0, "shrunk": [], "still_split": [], "skipped": True}

    cb = progress_cb or (lambda pct, msg: None)
    word = open_word()
    report = {"checked": 0, "shrunk": [], "still_split": []}
    try:
        doc = word.Documents.Open(os.path.abspath(path))
        doc.Repaginate()

        tables = [doc.Tables.Item(i + 1) for i in range(doc.Tables.Count)]
        for table in tables:
            if _is_rank_table(table) or _is_info_box(table):
                _fit_width(table)
        doc.Repaginate()

        targets = [(i + 1, t) for i, t in enumerate(tables) if _is_rank_table(t)]
        report["checked"] = len(targets)

        for idx, table in targets:
            title = _table_title(table) or f"table {idx}"
            passes = 0
            while passes < MAX_PASSES:
                start = _start_page_of(table)
                end = _end_page_of(table)
                if start == end:
                    cb(100, f"OK (1 page): {title[:52]}")
                    break
                if not _shrink_table(table):
                    report["still_split"].append(f"{title} (pages {start}-{end})")
                    cb(100, f"STILL SPLIT: {title[:52]}")
                    break
                doc.Repaginate()
                passes += 1
                report["shrunk"].append(f"{title}: pass {passes}")
            else:
                report["still_split"].append(f"{title} (max passes)")

        doc.Save()
        doc.Close(WD_SAVE_CHANGES)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    cb(100, "Word formatting finished")
    return report


def fit_bytes(blob, progress_cb=None):
    """Same pass on an in-memory .docx.  Returns the (possibly re-saved)
    bytes; returns the input untouched when Word is not available."""
    if not is_available() or not blob:
        return blob
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        with open(path, "wb") as f:
            f.write(blob)
        fit_college_tables(path, progress_cb=progress_cb)
        with open(path, "rb") as f:
            return f.read()
    except Exception as e:
        # The static layout is already JoSAA-correct - never fail a
        # download because the optional polish could not run.
        print(f"word_com.fit_bytes skipped: {e}")
        return blob
    finally:
        try:
            os.remove(path)
        except Exception:
            pass


# --------------------------------------------------------------------------- #
#  CLI
# --------------------------------------------------------------------------- #

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("usage: python word_com.py <college.docx>")
        raise SystemExit(2)
    if not is_available():
        print("Word formatting is not available on this machine "
              "(needs Windows + MS Word + pywin32).")
        raise SystemExit(1)

    def _cb(_pct, msg):
        if msg:
            print("  *", msg, flush=True)

    rep = fit_college_tables(sys.argv[1], progress_cb=_cb)
    print("tables checked :", rep["checked"])
    print("shrunk         :", len(rep["shrunk"]))
    for s in rep["shrunk"]:
        print("      ", s)
    print("still split    :", rep["still_split"])
