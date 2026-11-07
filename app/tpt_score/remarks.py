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


# --------------------------------------------------------------------------- remarks

def build_remarks(fs, limit: int = 6) -> str:
    """One short paragraph on what is wrong with this file, in plain words."""
    # Files that were zeroed before any check ran.
    zero = next((g for g in fs.gates if g.startswith(('G1', 'G2', 'G3'))), None)
    if zero:
        code = zero[:2]
        why = _gate_plain(zero)
        if code == 'G1':
            return 'Could not be opened (%s). Ask the GP to resend a valid, unprotected Excel file.' % fs.error
        if code == 'G2':
            return ('Does not look like a TPT file: %s. The header may be missing, the table '
                    'sideways, or the wrong file sent.' % why)
        return 'Nothing to score: %s.' % why

    parts: List[str] = []

    if fs.version != 'V7':
        parts.append('Wrong template version: this looks like %s, not V7.' % fs.version)

    if fs.missing_columns:
        names = [_pretty(n) for n in fs.missing_columns]
        parts.append('%d mandatory column%s missing: %s.'
                     % (len(names), '' if len(names) == 1 else 's', _listing(names, limit)))

    if fs.renamed_columns:
        pairs = ["'%s' should read '%s'" % (_pretty(a), _pretty(b)) for a, b in fs.renamed_columns]
        parts.append('%d column name%s differ%s slightly from the template: %s.'
                     % (len(pairs), '' if len(pairs) == 1 else 's', 's' if len(pairs) == 1 else '',
                        _listing(pairs, 4)))

    if fs.duplicate_columns:
        parts.append('Repeated column%s: %s.' % ('' if len(fs.duplicate_columns) == 1 else 's',
                                                _listing([_pretty(n) for n in fs.duplicate_columns], 4)))

    if fs.extra_columns:
        parts.append('%d extra column%s not in the template: %s.'
                     % (len(fs.extra_columns), '' if len(fs.extra_columns) == 1 else 's',
                        _listing([_pretty(n) for n in fs.extra_columns], 4)))

    if fs.header_cell and fs.header_cell != 'A1':
        parts.append('The table starts at %s instead of A1.' % fs.header_cell)

    s7 = fs.checks.get('S7')
    if s7 is not None and s7.n_affected:
        parts.append('%d row%s below the header are not positions (totals, notes or blank lines).'
                     % (s7.n_affected, '' if s7.n_affected == 1 else 's'))

    # Content problems: the biggest remaining point losses, except those already told above.
    told = {'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7'}
    content = [c for c in fs.top_fixes if c.id not in told and c.points_lost >= 0.5][:3]
    for c in content:
        detail = re.sub(r'\s*\(\d+ conditions? not evaluated\)', '', c.detail).rstrip('. ')
        parts.append('%s: %s.' % (c.title, detail))

    if any(g.startswith('G6') for g in fs.gates) and not parts:
        parts.append(_gate_plain(next(g for g in fs.gates if g.startswith('G6'))) + '.')

    return ' '.join(parts) if parts else 'No issues found.'
