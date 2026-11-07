"""Plain-language remarks for a scored file, and the gate reference table.

The remarks answer one question for a reader who is not a data person: what is
wrong with this file? They name the missing columns and the renamed ones.
"""
from __future__ import annotations

import re
from typing import List, Tuple

# --------------------------------------------------------------------------- gates

GATE_ORDER = ['G1', 'G2', 'G3', 'G4', 'G5', 'G6']


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


# --------------------------------------------------------------------------- findings

FINDING_LABELS = [('missing', 'Missing columns'), ('extra', 'Extra columns'),
                  ('renamed', 'Renamed columns'), ('blanks', 'Mandatory columns with blanks')]


def _clip(text, n: int = 50) -> str:
    text = _pretty(text)
    return text if len(text) <= n else text[:n - 1].rstrip() + '\u2026'


def column_findings(fs) -> dict:
    """Lists of short items for the four column-level findings.

    A file that is not recognisable as a TPT (gate G2) gets none: every column
    would be "missing" and the lists would only be noise.
    """
    if any(g.startswith('G2') for g in fs.gates):
        return {'missing': [], 'extra': [], 'renamed': [], 'blanks': []}
    missing = ['%s (mandatory)' % _pretty(n) for n in fs.missing_columns]
    missing += [_pretty(n) for n in fs.missing_optional]
    return {
        'missing': missing,
        'extra': [_clip(n) for n in fs.extra_columns],
        'renamed': ["%s -> %s" % (_clip(a), _pretty(b)) for a, b in fs.renamed_columns],
        'blanks': ['%s: %d of %d blank' % (_pretty(n), e, t) for n, e, t in fs.blank_mandatory],
    }


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
    """A short note on everything that is not a column-level finding."""
    zero = next((g for g in fs.gates if g.startswith(('G1', 'G2', 'G3'))), None)
    if zero:
        code = zero[:2]
        if code == 'G1':
            return 'Could not be opened. Ask the GP to resend a valid, unprotected Excel file.'
        if code == 'G2':
            return 'Not recognisable as a TPT file (%s).' % _gate_plain(zero)
        return 'Nothing to score: %s.' % _gate_plain(zero)

    parts: List[str] = []
    if fs.version != 'V7':
        parts.append('Template is %s, not V7' % fs.version)
    if fs.header_cell and fs.header_cell != 'A1':
        parts.append('table starts at %s, not A1' % fs.header_cell)
    s7 = fs.checks.get('S7')
    if s7 is not None and s7.n_affected:
        parts.append('%d non-position rows (totals, notes, blanks)' % s7.n_affected)
    if fs.duplicate_columns:
        parts.append('%d repeated columns' % len(fs.duplicate_columns))
    told = {'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'C1'}
    for c in [c for c in fs.top_fixes if c.id not in told and c.points_lost >= 0.5][:3]:
        parts.append(c.title[0].lower() + c.title[1:] if c.title[:2].isupper() is False else c.title)
    if not parts:
        findings = column_findings(fs)
        return '' if any(findings.values()) else 'No issues found.'
    text = '; '.join(parts)
    return text[0].upper() + text[1:] + '.'
