"""Which TPT template a submission was built on, judged from its column names only.

The column names of the official workbooks are compared (field 1000, the GP's own
version label, is ignored). What tells the versions apart, from the FinDatEx files
in templates/ (V6 14 Mar 2022, V7 25 Nov 2024, V8 post-consultation 8 Oct 2026):

* V7 adds fields 140-148 to V6; no field was renamed.
* V8 adds fields 150 and 151, and renames 62 and 148.

The workbooks are read when they sit next to the V7 template, so a newer FinDatEx
release changes the result without a code change; otherwise the values above are used.
"""
from __future__ import annotations

import glob
import os
import re
from typing import Dict, Optional

import openpyxl

from .rules import normalise, token_of

_NAME_RE = re.compile(r'^\s*(\d+)([a-z]?)_')

# Fallback when the V6 / V8 workbooks are not on disk (taken from them).
_FALLBACK = {
    'V6_missing': {'140', '141', '142', '143', '144', '145', '146', '147', '148'},
    'V8_only': {'150', '151'},
    'V8_renamed': {'62': '62_Conversion_factor', '148': '148_Economic_sector_NACE'},
}
_CACHE: Dict[str, dict] = {}


def _names(path: str) -> Dict[str, str]:
    """token -> column name, from the first column of a TPT definition workbook."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        out = {}
        for ws in wb.worksheets:
            for r in ws.iter_rows(values_only=True):
                v = r[0] if r else None
                if isinstance(v, str):
                    m = _NAME_RE.match(v)
                    if m:
                        out.setdefault(m.group(1) + m.group(2), v.strip())
        return out
    finally:
        wb.close()


def signatures(fields) -> dict:
    """What marks V6 and V8 against V7, read from the workbooks when available."""
    folder = os.path.dirname(os.path.abspath(getattr(fields, 'path', '') or '.'))
    if folder in _CACHE:
        return _CACHE[folder]
    sig = {k: (dict(v) if isinstance(v, dict) else set(v)) for k, v in _FALLBACK.items()}
    v7 = {f.token: f.name for f in fields.fields}

    def find(tag: str) -> Optional[str]:
        hits = [p for p in glob.glob(os.path.join(folder, '*.xlsx'))
                if re.search(r'V%s\b|V%s[ ._]' % (tag, tag), os.path.basename(p), re.I)]
        return sorted(hits)[0] if hits else None
    try:
        p6, p8 = find('6'), find('8')
        if p6:
            v6 = _names(p6)
            sig['V6_missing'] = {t for t in v7 if t not in v6}
        if p8:
            v8 = _names(p8)
            sig['V8_only'] = {t for t in v8 if t not in v7}
            sig['V8_renamed'] = {t: n for t, n in v8.items() if t in v7 and n != v7[t]}
    except Exception:
        pass                                  # unreadable workbook: keep the fallback
    _CACHE[folder] = sig
    return sig


def detect_version(header, fields, cmap) -> str:
    """'V6', 'V7', 'V8' or 'unknown', from the submission's column names."""
    sig = signatures(fields)
    tokens = {token_of(h) for h in header if h is not None}
    names = {normalise(h) for h in header if h is not None}
    v7 = {f.token: f.name for f in fields.fields}

    # V8: a column only V8 has, or more columns under their V8 name than their V7 name.
    if tokens & sig['V8_only']:
        return 'V8'
    v8_names = sum(1 for n in sig['V8_renamed'].values() if normalise(n) in names)
    v7_names = sum(1 for t in sig['V8_renamed'] if normalise(v7.get(t, '')) in names)
    if v8_names > v7_names:
        return 'V8'
    # V7: any of the columns V7 added to V6.
    if tokens & sig['V6_missing'] or set(cmap.mapped_tokens) & sig['V6_missing']:
        return 'V7'
    # Otherwise a full table without them is a V6 file.
    if cmap.n_mapped >= 60:
        return 'V6'
    return 'unknown'
