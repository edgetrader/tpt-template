"""S1 to S7: can the table be found and mapped onto the V7 template."""
from __future__ import annotations

from .base import CheckResult, Context, check, share

_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def _col_letter(n: int) -> str:
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = _ALPHABET[r] + s
    return s


@check('S1', 'Header sits in row 1, column A')
def s1(ctx: Context) -> CheckResult:
    t = ctx.table
    res = CheckResult('S1', 'Header sits in row 1, column A', 'structure')
    if t.at_origin:
        res.detail = 'header in A1'
        return res
    res.f = 1.0
    res.n_affected, res.n_applicable = 1, 1
    where = '%s%d' % (_col_letter(t.first_col), t.header_row)
    res.detail = 'header found at %s; the table was recovered from there' % where
    res.add(t.header_row, _col_letter(t.first_col), where, 'header in A1')
    return res


@check('S2', 'Template version is V7')
def s2(ctx: Context) -> CheckResult:
    res = CheckResult('S2', 'Template version is V7', 'structure')
    if ctx.version == 'V7':
        res.detail = 'V7'
        return res
    res.f = 1.0
    res.n_affected, res.n_applicable = 1, 1
    res.detail = 'template looks like %s, not V7' % ctx.version
    res.add(None, ctx.column_name('1000'), ctx.version, 'V7')
    return res


@check('S3', 'All mandatory columns present')
def s3(ctx: Context) -> CheckResult:
    res = CheckResult('S3', 'All mandatory columns present', 'structure')
    want = ctx.fields.mandatory_tokens
    missing = [t for t in want if not ctx.cmap.has(t)]
    res.n_applicable, res.n_affected = len(want), len(missing)
    res.f = share(len(missing), len(want))
    if missing:
        res.detail = '%d of %d mandatory columns missing' % (len(missing), len(want))
        for t in missing[:40]:
            f = ctx.fields.get(t)
            res.add(None, f.name if f else t, 'column absent', 'mandatory column present')
    else:
        res.detail = 'all %d present' % len(want)
    return res


@check('S4', 'Column names match the template')
def s4(ctx: Context) -> CheckResult:
    res = CheckResult('S4', 'Column names match the template', 'structure')
    cols = [c for c in ctx.cmap.columns if c.token is not None]
    renamed = [c for c in cols if c.renamed]
    res.n_applicable, res.n_affected = len(cols), len(renamed)
    res.f = share(len(renamed), len(cols))
    if renamed:
        res.detail = '%d of %d mapped columns do not carry the template name' % (len(renamed), len(cols))
        for c in renamed[:40]:
            f = ctx.fields.get(c.token)
            res.add(ctx.table.header_row, _col_letter(ctx.table.first_col + c.index),
                    c.header, f.name if f else c.token)
    else:
        res.detail = 'template names throughout'
    return res


@check('S5', 'No duplicate or ambiguous columns')
def s5(ctx: Context) -> CheckResult:
    res = CheckResult('S5', 'No duplicate or ambiguous columns', 'structure')
    cols = ctx.cmap.columns
    dups = ctx.cmap.duplicate_columns
    res.n_applicable, res.n_affected = len(cols), len(dups)
    res.f = share(len(dups), len(cols))
    if dups:
        res.detail = '%d columns repeat a field already mapped' % len(dups)
        for c in dups[:40]:
            res.add(ctx.table.header_row, _col_letter(ctx.table.first_col + c.index),
                    c.header, 'field %s already in column %s' %
                    (c.token, _col_letter(ctx.table.first_col + c.duplicate_of)))
    else:
        res.detail = 'no duplicates'
    return res


@check('S6', 'No unrecognised extra columns')
def s6(ctx: Context) -> CheckResult:
    res = CheckResult('S6', 'No unrecognised extra columns', 'structure')
    cols = ctx.cmap.columns
    unknown = [c for c in ctx.cmap.unknown_columns if str(c.header or '').strip()]
    blank = [c for c in ctx.cmap.unknown_columns if not str(c.header or '').strip()]
    affected = unknown + blank
    res.n_applicable, res.n_affected = len(cols), len(affected)
    res.f = share(len(affected), len(cols))
    if affected:
        res.detail = '%d columns are not template fields (%d blank headers)' % (len(affected), len(blank))
        for c in affected[:40]:
            res.add(ctx.table.header_row, _col_letter(ctx.table.first_col + c.index),
                    c.header if str(c.header or '').strip() else '(blank header)',
                    'a V7 template column')
    else:
        res.detail = 'every column is a template field'
    return res


@check('S7', 'Table body holds only positions')
def s7(ctx: Context) -> CheckResult:
    res = CheckResult('S7', 'Table body holds only positions', 'structure')
    t = ctx.table
    res.n_applicable, res.n_affected = t.rows_scanned, t.dropped_rows
    res.f = share(t.dropped_rows, t.rows_scanned)
    if t.dropped_rows:
        res.detail = '%d of %d rows below the header are not positions (blank, totals, notes or repeated headers)' % (
            t.dropped_rows, t.rows_scanned)
        for note in t.recovery:
            if note.startswith('row dropped') or 'blank rows removed' in note:
                res.add(None, 'table body', note, 'positions only below the header')
    else:
        res.detail = 'positions only'
    return res
