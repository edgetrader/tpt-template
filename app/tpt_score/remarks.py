"""Plain-language remarks for a scored file, and the gate reference table.

The remarks answer one question for a reader who is not a data person: what is
wrong with this file? They name the missing columns and the renamed ones.
"""
from __future__ import annotations

import re
from typing import List, Tuple

# --------------------------------------------------------------------------- gates

GATE_ORDER = ['G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7']


def gate_reference(cfg) -> List[Tuple[str, str, str, str]]:
    """Rows of (gate, what triggers it, effect on the score, plain meaning)."""
    g = cfg.gates
    mm = g.get('missing_mandatory', {})
    critical = sorted(cid for cid, spec in cfg.data.get('checks', {}).items()
                      if isinstance(spec, dict) and spec.get('critical'))
    return [
        ('G1', 'The file cannot be opened or read (corrupt, wrong format, password protected).',
         'Score 0', 'Nothing can be checked. Ask the GP to resend a valid Excel file.'),
        ('G2', 'No header row found, or fewer than %d columns match a TPT field.'
         % int(g.get('min_mapped_columns', 10)),
         'Score 0', 'This does not look like a TPT template, or the table is laid out sideways.'),
        ('G3', 'No position rows below the header, or the CIC column (field 12) cannot be found.',
         'Score 0', 'The template is empty, or the key column that says what each position is is missing.'),
        ('G4', 'The file is not on the V7 template (it looks like V6 or V8).',
         'Capped at %g' % float(g.get('wrong_version_cap', 40)),
         'Wrong template version. The columns may look similar, but it is not what was requested.'),
        ('G5', 'More than %d%% of the mandatory columns are missing.'
         % round(100 * float(mm.get('share', 0.20))),
         'Capped at %g' % float(mm.get('cap', 49)),
         'The file is a TPT in shape but too incomplete to use.'),
        ('G6', 'A critical check failed (%s).' % ', '.join(critical),
         'Capped at %g' % float(g.get('critical_failure_cap', 94)),
         'A serious flaw, so the file cannot reach Accept however good the rest is.'),
        ('G7', 'The workbook holds more than one TPT table (on other sheets, or a second header block on the same sheet).',
         'Capped at %g' % float(g.get('multiple_tables_cap', 40)),
         'Ambiguous submission: it is unclear which table is the real one. Ask the GP to send one table only.'),
    ]


# --------------------------------------------------------------------------- helpers

def _pretty(name) -> str:
    """'22_Quotation_currency' -> '22 Quotation currency'."""
    return re.sub(r'\s+', ' ', str(name).replace('_', ' ')).strip()


def _listing(items: List[str], limit: int) -> str:
    items = list(items)
    shown = ', '.join(items[:limit])
    if len(items) > limit:
        shown += ' and %d more' % (len(items) - limit)
    return shown


def _gate_plain(gate: str) -> str:
    """Strip the 'G4: ' prefix so the sentence reads naturally."""
    return gate.split(': ', 1)[1] if ': ' in gate else gate


# --------------------------------------------------------------------------- checks

CHECK_MEANING = {
    'S1': 'The table should start in cell A1 with the header row; title rows or an offset table are recovered but charged.',
    'S2': 'The file must be on the FinDatEx TPT V7 template, not V6 or V8.',
    'S3': 'Every column V7 marks mandatory (M) must be present.',
    'S4': 'Column headers should carry the V7 names; typos and short labels are matched but charged.',
    'S5': 'Each V7 field should appear once; repeated columns make it unclear which one to use.',
    'S6': 'Columns that are not V7 fields (comments, internal references) should be left out.',
    'S7': 'Below the header there should be positions only: no totals, notes, blank or repeated header rows.',
    'C1': 'Mandatory fields must hold a value on every position they apply to.',
    'C2': 'Conditional fields must be filled where their condition applies (for example the CIC category).',
    'C3': 'Every position needs an instrument identifier (field 14) and a CIC code (field 12).',
    'C4': 'Fund-level fields (fund ID, name, currency, NAV, dates) must be on every row.',
    'C5': 'Placeholders such as N/A, TBC, - or 0000 must not stand in for real values.',
    'V1': 'Numeric fields must hold numbers, not text or Excel errors such as #REF!.',
    'V2': 'Dates must be valid and written in one format, ideally yyyy-mm-dd.',
    'V3': 'Coded fields must use the allowed codes: CIC, ISO currency, ISO country, Y/N flags.',
    'V4': 'ISIN and LEI codes must pass their check digits.',
    'V5': 'Numbers must be in a sensible range and scale (for example weights as 0.25, not 25).',
    'V6': 'Text must be clean: no stray spaces, lost leading zeros or inconsistent case.',
    'K1': 'Fund-level fields must hold the same value on every row.',
    'K2': 'Valuation weights (field 26) must add up to 100%.',
    'K3': 'Market values in portfolio currency must add up to the fund NAV (field 5).',
    'K4': 'Quotation and portfolio-currency market values must agree through one FX rate per currency.',
    'K5': 'No position may appear twice.',
    'K6': 'The valuation date must not be after the reporting date, and the data must not be stale.',
}
DIMENSION_OF = {'S': 'Structure', 'C': 'Completeness', 'V': 'Validity and format', 'K': 'Consistency'}


def check_reference(cfg) -> List[Tuple[str, str, str, str, str, str]]:
    """Rows of (code, dimension, check, weight, critical, what it means)."""
    from .checks import CHECKS
    order = {'S': 0, 'C': 1, 'V': 2, 'K': 3}
    rows = []
    for cid in sorted(CHECKS, key=lambda c: (order.get(c[0], 9), int(c[1:]))):
        rows.append((cid, DIMENSION_OF.get(cid[0], ''), CHECKS[cid].title,
                     '%g' % cfg.weight(cid), 'Yes' if cfg.is_critical(cid) else '',
                     CHECK_MEANING.get(cid, '')))
    return rows


# --------------------------------------------------------------------------- findings

FINDING_LABELS = [('missing', 'Missing columns'), ('extra', 'Extra columns'),
                  ('renamed', 'Renamed columns'), ('blanks', 'Mandatory columns with blanks'),
                  ('types', 'Wrong data type'), ('dates', 'Date formats')]


def _clip(text, n: int = 50) -> str:
    text = _pretty(text)
    return text if len(text) <= n else text[:n - 1].rstrip() + '\u2026'


def _q(value, n: int = 24) -> str:
    """An example value, quoted and cut short."""
    v = ' '.join(str(value).split())
    return '"%s"' % (v if len(v) <= n else v[:n - 1].rstrip() + '\u2026')


def column_findings(fs) -> dict:
    """Lists of short items for the column-level findings.

    Every item reads "<column>: <what>" or names a column, so all lists look
    alike. A file that is not recognisable as a TPT (gate G2) gets none: every
    column would be "missing" and the lists would only be noise.
    """
    if any(g.startswith('G2') for g in fs.gates):
        return {k: [] for k, _ in FINDING_LABELS}
    missing = ['%s (mandatory)' % _pretty(n) for n in fs.missing_columns]
    missing += [_pretty(n) for n in fs.missing_optional]
    return {
        'missing': missing,
        'extra': [_clip(n) for n in fs.extra_columns],
        'renamed': ['%s \u2192 %s' % (_clip(a), _pretty(b)) for a, b in fs.renamed_columns],
        'blanks': ['%s: %d of %d blank' % (_pretty(n), e, t) for n, e, t in fs.blank_mandatory],
        'types': ['%s: %d of %d %s, e.g. %s' % (_pretty(n), bad, tot, why, _q(ex))
                  for n, exp, bad, tot, ex, why in fs.type_mismatches],
        'dates': [_date_item(c) for c in fs.date_formats],
    }


def _date_item(col) -> str:
    name, formats, invalid, populated, _standard = col
    parts = ['%s \u00d7%d' % (f, n) for f, n in formats]
    if invalid:
        parts.append('not a date \u00d7%d' % invalid)
    return '%s: %s' % (_pretty(name), ', '.join(parts))


def dates_standard(fs) -> bool:
    """True when every date column is written yyyy-mm-dd, the format V7 asks for."""
    return bool(fs.date_formats) and all(c[4] for c in fs.date_formats)


# What a failed check means, written as the problem rather than the rule.
CHECK_PROBLEM = {
    'C2': 'Conditional fields empty where needed',
    'C3': 'Some positions lack an ID or CIC code',
    'C4': 'Fund details missing on some rows',
    'C5': 'Placeholder values used (N/A, TBC)',
    'V4': 'Invalid ISIN or LEI codes',
    'V5': 'Numbers out of range',
    'V6': 'Untidy text (stray spaces, lost leading zeros)',
    'K1': 'Fund details differ between rows',
    'K2': 'Weights do not add up to 100%',
    'K3': 'Market values do not add up to the NAV',
    'K4': 'Exchange rates inconsistent',
    'K5': 'Duplicate positions',
    'K6': 'Dates out of order or out of date',
}
# Shown elsewhere as their own findings, so not repeated under "Other checks".
_TOLD = {'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'C1', 'V1', 'V2', 'V3'}

# Order and labels of every finding, as shown in the reports.
REMARK_LABELS = [('notscored', 'Not scored'), ('template', 'Template'),
                 ('tables', 'Extra TPT tables'), ('layout', 'Layout'),
                 ('repeated', 'Repeated columns'), ('checks', 'Other checks')]


def _detail(text: str) -> str:
    """A check's detail without the bookkeeping in brackets."""
    text = re.sub(r'\s*\((?:[^()]*(?:tolerance|not evaluated|tested)[^()]*)\)', '', text or '')
    return text.strip().rstrip('.')


def _cap(text: str) -> str:
    return text[:1].upper() + text[1:] if text else text


def remark_findings(fs) -> dict:
    """The findings that are not column lists, each a list of short items."""
    out = {k: [] for k, _ in REMARK_LABELS}
    zero = next((g for g in fs.gates if g.startswith(('G1', 'G2', 'G3'))), None)
    if zero:
        code = zero[:2]
        if code == 'G1':
            out['notscored'].append('File could not be opened (%s)' % (fs.error or 'unreadable'))
        elif code == 'G2':
            out['notscored'].append('Not recognisable as a TPT: %s' % _gate_plain(zero))
        else:
            out['notscored'].append(_cap(_gate_plain(zero)))
        return out
    if fs.version != 'V7':
        out['template'].append('%s template, not V7' % fs.version)
    out['tables'] = [_cap(w) for w in fs.extra_tables]
    if fs.header_cell and fs.header_cell != 'A1':
        out['layout'].append('Table starts at %s, not A1' % fs.header_cell)
    s7 = fs.checks.get('S7')
    if s7 is not None and s7.n_affected:
        out['layout'].append('%d non-position row%s (totals, notes or blanks)'
                             % (s7.n_affected, '' if s7.n_affected == 1 else 's'))
    out['repeated'] = [_clip(n) for n in fs.duplicate_columns]
    for c in sorted(fs.checks.values(), key=lambda c: -c.points_lost):
        if c.id in _TOLD or c.points_lost < 0.5:
            continue
        problem = CHECK_PROBLEM.get(c.id, c.title)
        detail = _detail(c.detail)
        out['checks'].append('%s: %s' % (problem, detail) if detail else problem)
    return out


def template_problems(fs) -> List[str]:
    """Template-level problems that are highlighted red: wrong version, more than one table."""
    f = remark_findings(fs)
    return f['template'] + (['%d extra TPT table%s' % (len(f['tables']), '' if len(f['tables']) == 1 else 's')]
                            if f['tables'] else [])


def format_finding(label: str, items: List[str], limit: int = 0) -> str:
    """'Label (n): item; item; ... and N more', the one format every finding uses."""
    shown = items if not limit else items[:limit]
    text = '%s (%d): %s' % (label, len(items), '; '.join(shown))
    if limit and len(items) > limit:
        text += ' and %d more' % (len(items) - limit)
    return text


def bullets(items: List[str], limit: int = 12) -> str:
    """One item per line; long lists are cut with a pointer to the full list."""
    items = list(items)
    if not items:
        return ''
    lines = ['\u2022 ' + i for i in items[:limit]]
    if len(items) > limit:
        lines.append('+ %d more (see Column findings)' % (len(items) - limit))
    return '\n'.join(lines)


def build_remarks(fs) -> str:
    """The non-column findings, one "Label (n): items" line each."""
    f = remark_findings(fs)
    lines = [format_finding(label, f[k]) for k, label in REMARK_LABELS if f[k]]
    if lines:
        return '\n'.join(lines)
    cols = column_findings(fs)
    problems = [k for k in ('missing', 'extra', 'renamed', 'blanks', 'types') if cols[k]]
    return '' if problems else 'No issues found.'
