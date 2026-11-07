"""Per-column profile of a file: data-type mismatches and date formats.

The checks (V1 to V5) charge points for bad values; this module answers the
reader's question instead: which columns hold data of the wrong kind, and in
what format are the dates written.
"""
from __future__ import annotations

import collections
from typing import Any, Dict, List, Tuple

from .rules import field_type
from .util import (ISO_COUNTRIES, ISO_CURRENCIES, as_date, as_number, is_blank, is_excel_error,
                   is_placeholder, text_of, valid_cic)

DATE_FORMAT_NAME = {
    'ISO': 'YYYY-MM-DD', 'ISO-slash': 'YYYY/MM/DD', 'compact': 'YYYYMMDD',
    'DMY-slash': 'DD/MM/YYYY', 'MDY-slash': 'MM/DD/YYYY',
    'DMY-dot': 'DD.MM.YYYY', 'MDY-dot': 'MM.DD.YYYY',
    'DMY-dash': 'DD-MM-YYYY', 'MDY-dash': 'MM-DD-YYYY',
    'excel': 'Excel date cell', 'serial': 'Excel serial number',
}
# What V7 asks for: ISO text, or a true Excel date.
STANDARD_DATE = {'ISO', 'excel'}

EXPECTED_LABEL = {
    'date': 'a date', 'integer': 'a whole number', 'number': 'a number', 'flag': 'Y or N',
    'currency': 'an ISO currency code', 'country': 'an ISO country code', 'cic': 'a CIC code',
}
CHECKED = set(EXPECTED_LABEL)
FLAG_TOKENS = {'11'}


def _kind_of(field, token: str) -> str:
    if token in FLAG_TOKENS:
        return 'flag'
    return field_type(field)


def _fits(kind: str, v) -> Tuple[bool, str]:
    """(value fits the expected type, short reason when it does not)."""
    if is_excel_error(v):
        return False, 'Excel error'
    if kind in ('number', 'integer'):
        num, from_text = as_number(v)
        if num is None:
            return False, 'not a number'
        if from_text:
            return False, 'number stored as text'
        if kind == 'integer' and abs(num - round(num)) > 1e-9:
            return False, 'not a whole number'
        return True, ''
    if kind == 'date':
        d, _ = as_date(v)
        return (d is not None), 'not a date'
    s = text_of(v).upper()
    if kind == 'cic':
        return valid_cic(v), 'not a CIC code'
    if kind == 'currency':
        return s in ISO_CURRENCIES, 'not a currency code'
    if kind == 'country':
        return s in ISO_COUNTRIES, 'not a country code'
    if kind == 'flag':
        return s in ('Y', 'N'), 'not Y or N'
    return True, ''


def profile_columns(ctx) -> Dict[str, List[Any]]:
    """Return {'type_mismatches': [...], 'date_formats': [...]}.

    type_mismatches: (column name, expected, bad cells, populated cells, example)
    date_formats:    (column name, [(format name, count)], invalid count, populated, standard?)
    """
    mismatches: List[Any] = []
    date_cols: List[Any] = []
    for t in ctx.cmap.mapped_tokens:
        f = ctx.fields.get(t)
        if f is None:
            continue
        kind = _kind_of(f, t)
        if kind not in CHECKED:
            continue
        populated = bad = 0
        example = ''
        reasons = collections.Counter()
        fmts = collections.Counter()
        for row in ctx.rows:
            v = ctx.value(row, t)
            if is_blank(v) or is_placeholder(v):
                continue
            populated += 1
            ok, why = _fits(kind, v)
            if not ok:
                bad += 1
                reasons[why] += 1
                if not example:
                    example = text_of(v)[:24]
            elif kind == 'date':
                fmts[as_date(v)[1]] += 1
        name = f.name
        if bad:
            mismatches.append((name, EXPECTED_LABEL[kind], bad, populated, example,
                               reasons.most_common(1)[0][0]))
        if kind == 'date' and populated:
            ordered = [(DATE_FORMAT_NAME.get(k, k), n, k) for k, n in fmts.most_common()]
            standard = bool(ordered) and len(ordered) == 1 and ordered[0][2] in STANDARD_DATE and not bad
            date_cols.append((name, [(a, b) for a, b, _ in ordered], bad, populated, standard))
    mismatches.sort(key=lambda x: -x[2])
    return {'type_mismatches': mismatches, 'date_formats': date_cols}
