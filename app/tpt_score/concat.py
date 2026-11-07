"""Combine the accepted TPT files of a run into one table.

    python tpt_concat.py --in <folder>      (same --in / settings as tpt_score.py)

Reads ``output/<input folder>/accepted_files.xlsx`` written by the last
tpt_score run, loads every file listed there (Accept and Minor fixes), maps its
columns onto the V7 template, converts every value to its column's type and
writes ``tpt_combined.xlsx`` beside the accepted list.

Columns: source (full path), gp_name, every V7 field in template order under
its template name, then any extra columns the GPs added, in first-seen order.

Types, by V7 field (from the template's codification):
  number / integer  real numbers; "1,234.50", "1.234,50", "(1,234)", "48.23%",
                    "1.5M" are converted; integer fields become whole numbers
  date              real Excel dates shown yyyy-mm-dd, from date cells, ISO,
                    dd/mm/yyyy, dd.mm.yyyy, yyyymmdd text or Excel serial numbers
  currency, country, CIC   trimmed, upper case
  Y/N flag (field 11)      Y or N (yes/no/true/false/1/0 accepted)
  text              trimmed, inner spaces collapsed; whole numbers read as
                    floats lose the trailing ".0"
A value that cannot be converted, an Excel error or a placeholder (N/A, TBC ...)
is left blank and listed on the "Conversion log" sheet with the original value.
Extra GP columns are kept as text, trimmed.
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
import warnings
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import __version__
from .loader import UnreadableFile, load_table
from .mapper import map_columns
from .rules import field_type, load_fieldset
from .util import as_date, as_number, is_excel_error, is_placeholder, text_of

COMBINED_NAME = 'tpt_combined.xlsx'
ACCEPTED_NAME = 'accepted_files.xlsx'
FLAG_TOKENS = {'11'}

HEAD_FILL = PatternFill('solid', fgColor='1F3864')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=10)
META_FILL = PatternFill('solid', fgColor='595959')
EXTRA_FILL = PatternFill('solid', fgColor='7F6000')


# --------------------------------------------------------------------------- conversion

class Bad(Exception):
    """A value that does not fit its column; the message says why."""


def _kind(field) -> str:
    if field.token in FLAG_TOKENS:
        return 'flag'
    return field_type(field)


def convert(kind: str, v: Any) -> Any:
    """The value in its column's type, None for an empty cell; raises Bad."""
    if v is None:
        return None
    if isinstance(v, str) and not v.strip():
        return None
    if is_excel_error(v):
        raise Bad('Excel error')
    if is_placeholder(v):
        raise Bad('placeholder')
    if kind in ('number', 'integer'):
        if isinstance(v, bool):
            raise Bad('not a number')
        num, _ = as_number(v)
        if num is None:
            raise Bad('not a number')
        if kind == 'integer':
            if abs(num - round(num)) > 1e-9:
                raise Bad('not a whole number')
            return int(round(num))
        return float(num)
    if kind == 'date':
        d, _ = as_date(v)
        if d is None:
            raise Bad('not a date')
        return d
    if kind in ('currency', 'country', 'cic'):
        return ' '.join(text_of(v).split()).upper()
    if kind == 'flag':
        s = text_of(v).strip().lower()
        if s in ('y', 'yes', 'true', '1', '1.0'):
            return 'Y'
        if s in ('n', 'no', 'false', '0', '0.0'):
            return 'N'
        raise Bad('not Y or N')
    return _text(v)


def _text(v: Any) -> Optional[str]:
    if isinstance(v, float) and v.is_integer():
        v = int(v)                                   # 12345.0 read from Excel -> "12345"
    if isinstance(v, dt.datetime):
        v = v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(sep=' ')
    elif isinstance(v, dt.date):
        v = v.isoformat()
    s = ' '.join(str(v).split())
    return s or None


# --------------------------------------------------------------------------- reading

def read_accepted(path: str) -> List[Tuple[str, str]]:
    """(full path, GP name) for every file on the accepted list."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb['Accepted files'] if 'Accepted files' in wb.sheetnames else wb.worksheets[0]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    for i, r in enumerate(rows):
        heads = [str(x or '').strip().lower() for x in r]
        if 'full path' in heads:
            p, g = heads.index('full path'), heads.index('gp name') if 'gp name' in heads else None
            out = []
            for row in rows[i + 1:]:
                if p < len(row) and row[p]:
                    gp = row[g] if g is not None and g < len(row) else ''
                    out.append((str(row[p]), str(gp or '')))
            return out
    raise ValueError('%s has no "Full path" column' % path)


def combine(files: List[Tuple[str, str]], fields):
    """Return (header, rows, log, file_notes)."""
    tokens = [f.token for f in fields.fields]
    kinds = {f.token: _kind(f) for f in fields.fields}
    extras: List[str] = []                 # extra column names, first seen first
    extra_seen = set()
    rows: List[Dict[str, Any]] = []
    log: List[List[Any]] = []
    notes: List[List[Any]] = []

    for path, gp in files:
        if not os.path.isfile(path):
            notes.append([path, gp, 0, 'file not found; skipped'])
            continue
        try:
            table = load_table(path, fields)
        except UnreadableFile as exc:
            notes.append([path, gp, 0, 'could not be read (%s); skipped' % exc])
            continue
        cmap = map_columns(table.header, fields)
        extra_cols = []
        for c in cmap.unknown_columns:
            name = ' '.join(str(c.header or '').split())
            if not name:
                continue
            extra_cols.append((c.index, name))
            if name.lower() not in extra_seen:
                extra_seen.add(name.lower())
                extras.append(name)
        bad = 0
        for i, raw in enumerate(table.rows):
            sheet_row = table.row_source[i] + 1 if i < len(table.row_source) else None
            out = {'source': path, 'gp_name': gp}
            for t in tokens:
                col = cmap.get(t)
                if col is None or col >= len(raw):
                    continue
                v = raw[col]
                try:
                    out[t] = convert(kinds[t], v)
                except Bad as why:
                    bad += 1
                    log.append([path, gp, sheet_row, fields.get(t).name, text_of(v)[:200], str(why)])
            for idx, name in extra_cols:
                if idx < len(raw):
                    out['x:' + name.lower()] = _text(raw[idx]) if raw[idx] is not None else None
            rows.append(out)
        notes.append([path, gp, len(table.rows),
                      'ok' if not bad else '%d value%s blanked (see Conversion log)'
                      % (bad, '' if bad == 1 else 's')])

    header = (['source', 'gp_name'] + [fields.get(t).name for t in tokens] + extras)
    keys = ['source', 'gp_name'] + tokens + ['x:' + e.lower() for e in extras]
    table_rows = [[r.get(k) for k in keys] for r in rows]
    meta = {'n_v7': len(tokens), 'n_extra': len(extras),
            'kinds': [None, None] + [kinds[t] for t in tokens] + ['text'] * len(extras)}
    return header, table_rows, log, notes, meta


# --------------------------------------------------------------------------- writing

def write_combined(path: str, header, rows, log, notes, meta, source_list: str) -> str:
    from .reporters.excel import save          # same Excel-safety pass as the reports
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Combined'
    ws.append(header)
    for j, name in enumerate(header, start=1):
        c = ws.cell(1, j)
        c.font = HEAD_FONT
        c.fill = META_FILL if j <= 2 else (HEAD_FILL if j <= 2 + meta['n_v7'] else EXTRA_FILL)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        ws.column_dimensions[get_column_letter(j)].width = 70 if j == 1 else (28 if j == 2 else 18)
    date_cols = [j for j, k in enumerate(meta['kinds'], start=1) if k == 'date']
    for r in rows:
        ws.append(r)
    for j in date_cols:
        letter = get_column_letter(j)
        for cell in ws[letter][1:]:
            if cell.value is not None:
                cell.number_format = 'yyyy-mm-dd'
    ws.freeze_panes = 'C2'
    ws.row_dimensions[1].height = 60
    if rows:
        ws.auto_filter.ref = 'A1:%s%d' % (get_column_letter(len(header)), len(rows) + 1)

    lg = wb.create_sheet('Conversion log')
    lg.append(['Source', 'GP name', 'Row in file', 'Column', 'Original value', 'Why it was blanked'])
    for r in log:
        lg.append(r)
    for j, w in enumerate([70, 28, 11, 50, 40, 24], start=1):
        lg.column_dimensions[get_column_letter(j)].width = w
        lg.cell(1, j).font, lg.cell(1, j).fill = HEAD_FONT, HEAD_FILL
    lg.freeze_panes = 'A2'
    if log:
        lg.auto_filter.ref = 'A1:F%d' % (len(log) + 1)

    fl = wb.create_sheet('Files')
    fl.append(['Source', 'GP name', 'Rows', 'Note'])
    for r in notes:
        fl.append(r)
    for j, w in enumerate([70, 28, 8, 50], start=1):
        fl.column_dimensions[get_column_letter(j)].width = w
        fl.cell(1, j).font, fl.cell(1, j).fill = HEAD_FONT, HEAD_FILL
    fl.append([])
    fl.append(['Taken from %s' % source_list])
    fl.append(['Grey columns: source and GP. Blue: V7 fields in template order. '
               'Brown: extra columns the GPs added.'])
    return save(wb, path)


# --------------------------------------------------------------------------- command line

def build_parser() -> argparse.ArgumentParser:
    from .cli import DEFAULT_OUT
    p = argparse.ArgumentParser(
        prog='tpt_concat',
        description='Combine the accepted TPT files of the last tpt_score run over a folder '
                    'into one table, tpt_combined.xlsx, in the same output folder.')
    p.add_argument('-i', '--in', dest='input', default=None, metavar='FOLDER',
                   help='the input folder that was scored (as given to tpt_score.py)')
    p.add_argument('-s', '--settings', default=None, metavar='FILE',
                   help='settings file (default: config/settings.yaml)')
    p.add_argument('-o', '--out', dest='output', default=DEFAULT_OUT, metavar='FOLDER',
                   help='the output root used by tpt_score.py (default: %s)' % DEFAULT_OUT)
    p.add_argument('-t', '--template', metavar='FILE', help='the FinDatEx V7 workbook')
    p.add_argument('-c', '--config', metavar='FILE', help='rules.yaml')
    p.add_argument('-q', '--quiet', action='store_true')
    p.add_argument('--version', action='version', version='tpt_concat %s' % __version__)
    return p


def run(args=None) -> int:
    warnings.filterwarnings('ignore')
    from .cli import input_label, resolve_template
    from .config import load_config
    from .settings import SettingsError, load_settings
    opts = build_parser().parse_args(args)
    try:
        settings = load_settings(opts.settings)
    except SettingsError as exc:
        print('settings error: %s' % exc, file=sys.stderr)
        return 2
    given = opts.input
    if not opts.input:
        opts.input = settings.input_folder
    elif settings.input_folder and not os.path.isabs(os.path.expanduser(opts.input)):
        opts.input = os.path.normpath(os.path.join(settings.input_folder, opts.input))
    if not opts.input:
        print('no input folder: pass --in FOLDER or set input_folder in config/settings.yaml',
              file=sys.stderr)
        return 2

    outdir = os.path.join(opts.output, input_label(opts.input, given))
    accepted = os.path.join(outdir, ACCEPTED_NAME)
    if not os.path.isfile(accepted):
        print('no %s in %s: run tpt_score.py on this folder first' % (ACCEPTED_NAME, outdir),
              file=sys.stderr)
        return 2
    try:
        cfg = load_config(opts.config)
        template = resolve_template(cfg, opts.template, opts.input)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 2
    fields = load_fieldset(template)

    files = read_accepted(accepted)
    if not opts.quiet:
        print('tpt_concat %s | %d accepted files from %s' % (__version__, len(files), accepted))
    header, rows, log, notes, meta = combine(files, fields)
    path = write_combined(os.path.join(outdir, COMBINED_NAME), header, rows, log, notes, meta,
                          accepted)
    skipped = sum(1 for n in notes if n[2] == 0 and 'skipped' in n[3])
    print('%d rows from %d files (%d skipped), %d V7 columns + %d extra, %d values blanked'
          % (len(rows), len(files) - skipped, skipped, meta['n_v7'], meta['n_extra'], len(log)))
    print('  combined: %s' % path)
    return 0
