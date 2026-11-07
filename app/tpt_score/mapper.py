"""Map the columns of a submission onto V7 field numbers.

Three passes, in order of confidence:

1. the leading field number (``22_Market_valuation...`` and ``8b_...``),
2. the exact template name with the number stripped,
3. a fuzzy match on the name, plus a table of short labels GPs commonly use.

Anything still unmatched is an unknown column.
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field as dc_field
from typing import Any, Dict, List, Optional

from .rules import FieldSet, normalise, token_of

# Short labels seen in GP deliveries, mapped to the field they mean.
SHORT_LABELS = {
    'fundid': '1', 'portfolioid': '1', 'fundidentifier': '1',
    'idtype': '2', 'typeofid': '2',
    'fundname': '3', 'portfolioname': '3',
    'fundccy': '4', 'fundcurrency': '4', 'portfolioccy': '4',
    'nav': '5', 'netassetvalue': '5',
    'valuationdate': '6', 'valdate': '6', 'navdate': '6',
    'reportingdate': '7', 'repdate': '7', 'asofdate': '7',
    'shareprice': '8', 'cashratio': '9', 'cash': '9',
    'scrdelivered': '11', 'completescrdelivery': '11',
    'cic': '12', 'ciccode': '12',
    'isin': '14', 'instrumentid': '14', 'identificationcode': '14', 'isinid': '14',
    'idtypeinstr': '15', 'instrumentidtype': '15',
    'instrument': '17', 'instrumentname': '17', 'assetname': '17',
    'qty': '18', 'quantity': '18', 'nominal': '19', 'nominalvalue': '19',
    'ccy': '21', 'quotationcurrency': '21', 'currency': '21',
    'mvqc': '22', 'marketvalueqc': '22', 'marketvalue': '22',
    'mvpc': '24', 'marketvaluepc': '24',
    'weight': '26', 'valuationweight': '26',
    'issuer': '46', 'issuername': '46', 'issuerlei': '47', 'lei': '47',
    'issuercountry': '52', 'country': '52',
    'tptversion': '1000', 'version': '1000',
}


@dataclass
class Column:
    index: int
    header: Any
    token: Optional[str] = None
    how: str = 'unknown'          # number, name, fuzzy, alias, unknown
    duplicate_of: Optional[int] = None
    renamed: bool = False         # matched a field but the text is not the template's


@dataclass
class ColumnMap:
    columns: List[Column] = dc_field(default_factory=list)
    index: Dict[str, int] = dc_field(default_factory=dict)   # token -> column index

    def get(self, tokenid) -> Optional[int]:
        return self.index.get(tokenid)

    def has(self, tokenid) -> bool:
        return tokenid in self.index

    def value(self, row, tokenid):
        i = self.index.get(tokenid)
        if i is None or i >= len(row):
            return None
        return row[i]

    @property
    def mapped_tokens(self) -> List[str]:
        return list(self.index.keys())

    @property
    def n_mapped(self) -> int:
        return len(self.index)

    @property
    def unknown_columns(self) -> List[Column]:
        return [c for c in self.columns if c.token is None]

    @property
    def duplicate_columns(self) -> List[Column]:
        return [c for c in self.columns if c.duplicate_of is not None]

    @property
    def renamed_columns(self) -> List[Column]:
        return [c for c in self.columns if c.token is not None and c.renamed]


def _strip_suffix_noise(text: str) -> str:
    """Drop pandas-style and GP-style suffixes: ``22_... .1``, ``(copy)``."""
    s = re.sub(r'\.\d+\s*$', '', str(text))
    s = re.sub(r'\s*\((copy|dup|duplicate|2)\)\s*$', '', s, flags=re.I)
    return s


def map_columns(header: List[Any], fields: FieldSet, fuzzy_cutoff: float = 0.86) -> ColumnMap:
    cmap = ColumnMap()
    names = list(fields.by_name.keys())

    for i, h in enumerate(header):
        col = Column(index=i, header=h)
        text = '' if h is None else str(h).strip()
        if text:
            # A template name may itself end in a number (148_..._NACE2.1), so
            # the untouched text is tried before the suffix is stripped.
            raw = text if normalise(text) in fields.by_name else _strip_suffix_noise(text)
            norm = normalise(raw)

            tok = token_of(raw)
            if tok and tok in fields:
                col.token, col.how = tok, 'number'
                col.renamed = norm != normalise(fields.get(tok).name)
            elif norm and norm in fields.by_name:
                col.token, col.how = fields.by_name[norm].token, 'name'
                col.renamed = True
            elif norm in SHORT_LABELS and SHORT_LABELS[norm] in fields:
                col.token, col.how = SHORT_LABELS[norm], 'alias'
                col.renamed = True
            elif norm:
                near = difflib.get_close_matches(norm, names, n=1, cutoff=fuzzy_cutoff)
                if near:
                    col.token, col.how = fields.by_name[near[0]].token, 'fuzzy'
                    col.renamed = True
        cmap.columns.append(col)

    # First occurrence wins; later ones are recorded as duplicates so that a
    # second column cannot silently overwrite the first.
    for col in cmap.columns:
        if col.token is None:
            continue
        if col.token in cmap.index:
            col.duplicate_of = cmap.index[col.token]
        else:
            cmap.index[col.token] = col.index
    return cmap


def detect_version(header: List[Any], rows, cmap: ColumnMap, fields: FieldSet) -> str:
    """Return 'V7', 'V6', 'V8' or 'unknown' for the submission."""
    # The version label in field 1000 is the GP's own statement.
    stated = ''
    i = cmap.get('1000')
    if i is not None:
        for r in rows[:5]:
            if i < len(r) and r[i] is not None:
                stated = str(r[i]).strip()
                break

    mapped = set(cmap.mapped_tokens)
    v7_only = {t for t in ('140', '141', '142', '143', '144', '145', '146', '147', '148') if t in fields}
    has_v7 = bool(v7_only & mapped)
    # Fields introduced after V7; present only in a newer template.
    has_newer = bool({'150', '151', '152', '153'} & mapped)

    m = re.search(r'v\s*([0-9]+)', stated, re.I)
    if m:
        label = 'V' + m.group(1)
        if label in ('V6', 'V7', 'V8'):
            # The label is believed only when the columns do not contradict it.
            if label == 'V7' and not has_v7 and not has_newer:
                return 'V6'
            if label == 'V7' and has_newer:
                return 'V8'
            return label

    if has_newer:
        return 'V8'
    if has_v7:
        return 'V7'
    if cmap.n_mapped >= 60:
        return 'V6'
    return 'unknown'
