"""Open a GP submission and recover the position table from it.

GP files rarely arrive as a clean table in A1: there may be title rows, the
table may start at C4, there may be a TOTAL row at the bottom or blank rows in
the middle. The loader finds the table, keeps only position rows, and records
everything it had to repair in a recovery log so the structure checks can
charge for it.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field as dc_field
from typing import Any, List, Tuple

import openpyxl

from .rules import FieldSet, normalise, token_of

READABLE_EXT = {'.xlsx', '.xlsm', '.xltx', '.xltm'}

# How far into the sheet to look for formulas. Reading the whole workbook a
# second time doubles the time spent on every file for a note, so the scan
# stops once the pattern would have shown itself.
FORMULA_SCAN_ROWS = 60

# A row naming at least this many template fields is a table header.
MIN_TABLE_COLUMNS = 10

# Rows that are plainly not positions.
_TOTAL_RE = re.compile(r'^\s*(total|totals|sum|subtotal|grand\s*total|sous[- ]total|somme)\b', re.I)
_NOTE_RE = re.compile(r'^\s*(prepared|source|note|notes|disclaimer|contact|checked|reviewed|generated)\b', re.I)
_DIVIDER_RE = re.compile(r'^\s*[-=_*\s]{3,}|^\s*---')


class UnreadableFile(Exception):
    """The file cannot be opened at all (gate G1)."""


@dataclass
class Table:
    header: List[Any]
    rows: List[List[Any]]
    sheet: str
    header_row: int = 1          # 1-based position in the sheet
    first_col: int = 1           # 1-based
    recovery: List[str] = dc_field(default_factory=list)
    dropped_rows: int = 0        # non-position rows removed from the body
    rows_scanned: int = 0        # body rows seen before dropping
    sheets_available: List[str] = dc_field(default_factory=list)
    hidden_rows: int = 0
    hidden_cols: int = 0
    merged_cells: int = 0
    formula_cells: int = 0
    leading_blank_cols: int = 0
    trailing_blank_cols: int = 0
    extra_tables: List[str] = dc_field(default_factory=list)   # other TPT tables in the workbook

    @property
    def at_origin(self) -> bool:
        return self.header_row == 1 and self.first_col == 1

    def note(self, text: str):
        self.recovery.append(text)


def _cell_text(v) -> str:
    return '' if v is None else str(v).strip()


def _row_blank(row) -> bool:
    return all(_cell_text(v) == '' for v in row)


def _has_blank_cells(grid, scan_rows: int = 200) -> bool:
    """Whether any cell inside the used range came back empty. A grid with no
    gaps cannot be hiding an uncached formula, so the second read is skipped."""
    for row in grid[:scan_rows]:
        if any(v is None for v in row):
            return True
    return False


def _score_header(row, fields: FieldSet) -> Tuple[int, int]:
    """How template-like a row looks: (tokens recognised, first column used).

    The left edge of the table is the first non-empty cell of the header row,
    not the first cell the template recognises: a column the GP renamed sits
    left of the recognised ones often enough, and dropping it would lose data.
    """
    hits, first = 0, None
    for i, v in enumerate(row):
        t = _cell_text(v)
        if not t:
            continue
        if first is None:
            first = i
        tok = token_of(t)
        if (tok is not None and tok in fields) or normalise(t) in fields.by_name:
            hits += 1
    return hits, (0 if first is None else first)


def _pick_sheet(wb, fields: FieldSet):
    """The sheet whose best row looks most like the template header.

    A sheet called "Data" is preferred only among sheets that actually look
    like the template: a workbook whose cover page happens to be named "Data"
    must not hide the positions on another sheet.
    """
    names = wb.sheetnames
    hits_of = {}
    for n in names:
        ws = wb[n]
        best_row = 0
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 40), values_only=True):
            h, _ = _score_header(list(row), fields)
            best_row = max(best_row, h)
        hits_of[n] = best_row

    top = max(hits_of.values()) if hits_of else 0
    # Sheets within reach of the best one are the candidates; "Data" wins there.
    plausible = [n for n in names if hits_of[n] >= max(10, top * 0.9)]
    chosen = None
    for n in plausible:
        if n.strip().lower() == 'data':
            chosen = n
            break
    if chosen is None:
        chosen = max(names, key=lambda n: hits_of[n]) if names else None

    others = [n for n in names if n != chosen and hits_of[n] >= MIN_TABLE_COLUMNS]
    note = ''
    if len(names) > 1 and chosen:
        note = 'sheet "%s" chosen (%d template columns recognised there)' % (
            chosen, hits_of.get(chosen, 0))
    return wb[chosen], chosen, note, others


def load_table(path: str, fields: FieldSet, max_header_scan: int = 40) -> Table:
    """Read ``path`` and return the recovered position table."""
    ext = os.path.splitext(path)[1].lower()
    if ext not in READABLE_EXT:
        raise UnreadableFile('not an Excel workbook (%s)' % (ext or 'no extension'))
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:                      # corrupt, encrypted, wrong format
        raise UnreadableFile('%s: %s' % (type(exc).__name__, str(exc)[:120]))

    try:
        ws, sheet_name, sheet_note, other_sheets = _pick_sheet(wb, fields)
        grid = [list(r) for r in ws.iter_rows(values_only=True)]

        hidden_rows = sum(1 for d in ws.row_dimensions.values() if d.hidden)
        hidden_cols = sum(1 for d in ws.column_dimensions.values() if d.hidden)
        merged = len(ws.merged_cells.ranges)
        sheets_available = list(wb.sheetnames)
    finally:
        wb.close()

    # Formula cells read as None with data_only=True when Excel never cached a
    # value. Opening the workbook a second time to count them doubles the time
    # spent on every file, so it is only done when something actually came back
    # empty, and then in read-only mode.
    formula_cells = 0
    if _has_blank_cells(grid):
        try:
            wb2 = openpyxl.load_workbook(path, data_only=False, read_only=True)
            ws2 = wb2[sheet_name]
            # Whether the file is built on formulas shows in its first rows; the
            # exact count is not worth reading the whole workbook again.
            for row in ws2.iter_rows(max_row=FORMULA_SCAN_ROWS, values_only=True):
                for v in row:
                    if isinstance(v, str) and v.startswith('='):
                        formula_cells += 1
            wb2.close()
        except Exception:
            pass

    table = Table(header=[], rows=[], sheet=sheet_name,
                  sheets_available=sheets_available, hidden_rows=hidden_rows,
                  hidden_cols=hidden_cols, merged_cells=merged,
                  formula_cells=formula_cells)
    if sheet_note:
        table.note(sheet_note)
    for other in other_sheets:
        table.extra_tables.append('sheet "%s"' % other)
        table.note('a second TPT table sits on sheet "%s"' % other)
    if len(sheets_available) > 1:
        table.note('workbook has %d sheets: %s' % (len(sheets_available), ', '.join(sheets_available)))
    if hidden_rows or hidden_cols:
        table.note('%d hidden rows and %d hidden columns were read' % (hidden_rows, hidden_cols))
    if merged:
        table.note('%d merged cell ranges on the sheet' % merged)
    if formula_cells:
        table.note('formulas found in the first %d rows (%d cells); without a cached '
                   'value they read as empty' % (FORMULA_SCAN_ROWS, formula_cells))

    if not grid:
        return table

    # --- find the header row ---------------------------------------------
    best_i, best_hits, best_first = None, 0, 0
    for i, row in enumerate(grid[:max_header_scan]):
        hits, first = _score_header(row, fields)
        if hits > best_hits:
            best_i, best_hits, best_first = i, hits, first
    if best_i is None:
        return table                                   # no header: gate G2

    table.header_row = best_i + 1
    table.first_col = best_first + 1
    if best_i > 0:
        table.note('header found in row %d, not row 1 (%d rows above it)' % (best_i + 1, best_i))
    if best_first > 0:
        table.note('table starts in column %s, not column A' % _col_letter(best_first + 1))
        table.leading_blank_cols = best_first

    header = grid[best_i][best_first:]
    # Trim trailing columns that are empty in the header and in every row.
    body_raw = [r[best_first:] for r in grid[best_i + 1:]]
    width = len(header)
    while width > 0 and _cell_text(header[width - 1]) == '' and \
            all(width > len(r) or _cell_text(r[width - 1]) == '' for r in body_raw):
        width -= 1
    if width < len(header):
        table.trailing_blank_cols = len(header) - width
        table.note('%d empty trailing columns ignored' % table.trailing_blank_cols)
    header = header[:width]
    table.header = header

    # Excel reports a used range that often runs far past the last position, so
    # rows after the last one holding anything are not part of the table at all
    # and must not be charged to the GP as "not positions".
    while body_raw and _row_blank(body_raw[-1][:width]):
        body_raw.pop()

    # --- keep only position rows ------------------------------------------
    header_norm = [normalise(h) for h in header]
    body: List[List[Any]] = []
    blank_runs = 0
    after_blank = False
    for idx, r in enumerate(body_raw):
        row = list(r[:width]) + [None] * max(0, width - len(r))
        table.rows_scanned += 1
        first_text = _cell_text(row[0])
        joined = ' '.join(_cell_text(v) for v in row[:4])

        if _row_blank(row):
            blank_runs += 1
            table.dropped_rows += 1
            after_blank = True
            continue
        was_after_blank, after_blank = after_blank, False
        # A header block that starts after a gap is a second table on this sheet.
        # A header repeated straight under data is only a page-break repeat.
        if was_after_blank and sum(1 for v in row if _cell_text(v)) >= MIN_TABLE_COLUMNS and \
                _score_header(row, fields)[0] >= MIN_TABLE_COLUMNS:
            where = 'sheet "%s" row %d' % (sheet_name, best_i + 2 + idx)
            table.extra_tables.append(where)
            table.dropped_rows += 1
            table.note('a second TPT table header at %s' % where)
            continue
        if _TOTAL_RE.match(first_text) or _TOTAL_RE.match(joined):
            table.dropped_rows += 1
            table.note('row dropped as a totals row: %r' % first_text[:40])
            continue
        if _NOTE_RE.match(first_text) or _DIVIDER_RE.match(first_text):
            table.dropped_rows += 1
            table.note('row dropped as a note or divider: %r' % first_text[:40])
            continue
        if [normalise(v) for v in row] == header_norm:
            table.dropped_rows += 1
            table.note('row dropped as a repeated header')
            continue
        # A row with one value and nothing else is a section label, not a position.
        if sum(1 for v in row if _cell_text(v)) == 1 and len(first_text) > 0 and \
                not any(_cell_text(v) for v in row[1:]):
            table.dropped_rows += 1
            table.note('row dropped as a section label: %r' % first_text[:40])
            continue
        body.append(row)

    if blank_runs:
        table.note('%d blank rows removed from the table body' % blank_runs)
    table.rows = body
    return table


def _col_letter(n: int) -> str:
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s
