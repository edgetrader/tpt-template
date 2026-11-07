"""Field definitions and applicability rules, read from the FinDatEx TPT V7 workbook.

The workbook is the single source of truth: column A holds the field name
(``22_Market_valuation_in_quotation_currency_(A)``), column K holds the
Mandatory / Conditional / Optional / Indicative / N-A flag, and columns L to AA
hold the applicability per CIC category, either ``x`` or a conditional text such
as ``x if item 120 set to 1``.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field as dc_field
from typing import Dict, List, Optional

import openpyxl

# Column layout of the template sheet.
COL_NAME = 1
COL_DEFINITION = 3
COL_CODIFICATION = 4
COL_FLAG = 11
COL_CIC_FIRST = 12
COL_CIC_LAST = 27

# CIC categories in the order of columns L to AA.
CIC_CATEGORIES = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
                  'A', 'B', 'C', 'D', 'E', 'F']

TOKEN_RE = re.compile(r'^\s*(\d+)\s*([a-zA-Z]?)\s*[_\s(]')


def token_of(name) -> Optional[str]:
    """``8b_Total_number_of_shares`` -> ``8b``. None when there is no number."""
    if name is None:
        return None
    text = str(name)
    m = TOKEN_RE.match(text)
    if not m:
        m = re.match(r'^\s*(\d+)\s*([a-zA-Z]?)\s*$', text)
    if not m:
        # GP numbering styles: "Field 26 - Valuation weight", "124) Duration",
        # "123-Fund CIC", "94bis_Vega" (bis = the lettered field 94b).
        m = re.match(r'^\s*(?:field|item|no\.?|#)?\s*(\d{1,4})\s*(bis|[a-zA-Z](?![a-zA-Z]))?\s*[)\-:._ ]',
                     text, re.I)
        if m:
            letter = (m.group(2) or '').lower()
            return m.group(1) + ('b' if letter == 'bis' else letter)
        return None
    return m.group(1) + m.group(2).lower()


def normalise(text) -> str:
    """Lower-case alphanumeric form of a field name, for fuzzy matching."""
    if text is None:
        return ''
    s = str(text)
    s = TOKEN_RE.sub('', s, count=1) if TOKEN_RE.match(s) else s
    return re.sub(r'[^a-z0-9]+', '', s.lower())


@dataclass
class Field:
    token: str
    number: int
    letter: str
    name: str
    flag: str                       # M, C, O, I, N/A or ''
    cic: Dict[str, Optional[str]] = dc_field(default_factory=dict)
    definition: str = ''
    codification: str = ''

    @property
    def mandatory(self) -> bool:
        return self.flag == 'M'

    @property
    def conditional(self) -> bool:
        return self.flag == 'C'

    def applies_to_cic(self, cic_category: Optional[str]):
        """True / False / the conditional text for this CIC category.

        Returns True when the cell holds a plain ``x``, False when it is empty,
        and the text itself when the applicability is conditional. An unknown
        category gives None, meaning "cannot be decided".
        """
        if cic_category is None or cic_category not in self.cic:
            return None
        v = self.cic.get(cic_category)
        if v is None:
            return False
        s = str(v).strip()
        if not s:
            return False
        if s.lower() == 'x':
            return True
        return s


class FieldSet:
    """All fields of the template, addressable by token."""

    def __init__(self, fields: List[Field], version_label: str = ''):
        self.fields = fields
        self.version_label = version_label
        self.by_token: Dict[str, Field] = {f.token: f for f in fields}
        self.by_name: Dict[str, Field] = {normalise(f.name): f for f in fields}

    def __len__(self):
        return len(self.fields)

    def __contains__(self, tokenid):
        return tokenid in self.by_token

    def get(self, tokenid) -> Optional[Field]:
        return self.by_token.get(tokenid)

    @property
    def tokens(self) -> List[str]:
        return [f.token for f in self.fields]

    @property
    def mandatory_tokens(self) -> List[str]:
        return [f.token for f in self.fields if f.mandatory]

    @property
    def conditional_tokens(self) -> List[str]:
        return [f.token for f in self.fields if f.conditional]


def load_fieldset(template_path: str) -> FieldSet:
    """Read the field definitions out of the V7 template workbook."""
    wb = openpyxl.load_workbook(template_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()

    version_label = ''
    for r in rows[:8]:
        for v in r:
            if v and str(v).lower().startswith('version'):
                version_label = str(v).strip()
                break
        if version_label:
            break

    # The header row is the one holding NUM_DATA in column A.
    header_row = 0
    for i, r in enumerate(rows):
        if r and str(r[0]).strip().upper() == 'NUM_DATA':
            header_row = i
            break

    fields: List[Field] = []
    for r in rows[header_row + 1:]:
        if not r or r[0] is None:
            continue
        name = str(r[0]).strip()
        tok = token_of(name)
        if tok is None:
            continue                      # section banner such as "Portfolio Characteristics"
        m = re.match(r'^(\d+)([a-zA-Z]?)$', tok)
        number, letter = int(m.group(1)), m.group(2)

        def cell(idx):
            return r[idx - 1] if len(r) >= idx else None

        flag = str(cell(COL_FLAG) or '').strip().upper()
        cic = {}
        for k, c in enumerate(range(COL_CIC_FIRST, COL_CIC_LAST + 1)):
            v = cell(c)
            cic[CIC_CATEGORIES[k]] = None if v is None else str(v).strip()
        fields.append(Field(
            token=tok, number=number, letter=letter, name=name, flag=flag, cic=cic,
            definition=str(cell(COL_DEFINITION) or '').strip(),
            codification=str(cell(COL_CODIFICATION) or '').strip(),
        ))
    return FieldSet(fields, version_label)


# --- conditional applicability -------------------------------------------------
# Only the machine-readable shapes are evaluated. Anything else is reported as
# "not evaluated" and carries no penalty, as agreed in the proposal.

_COND_ITEM_SET = re.compile(r'if\s+item\s+(\d+[a-z]?)\s+set\s+to\s+"?([^"\s]+)"?', re.I)
_COND_ITEM_NOT_BLANK = re.compile(r'if\s+item\s+(\d+[a-z]?)\s+is\s+not\s+blank', re.I)
_COND_CIC_LIST = re.compile(r'^x?\s*for\s+([0-9A-F]{1,2}(?:\s*,\s*[0-9A-F]{1,2})*)\s*,?\s*$', re.I)


def evaluate_condition(text: str, row_values, colmap, cic_code: Optional[str]):
    """Decide whether a conditional field applies to one row.

    Returns True, False, or None when the condition cannot be read.
    ``row_values`` is the row, ``colmap`` maps token -> column index.
    """
    s = ' '.join(str(text).split())

    m = _COND_ITEM_SET.search(s)
    if m:
        tok, want = m.group(1).lower(), m.group(2).strip().strip('"')
        idx = colmap.get(tok)
        if idx is None or idx >= len(row_values):
            return None
        got = row_values[idx]
        if got is None:
            return False
        got = str(got).strip().strip('"')
        if re.fullmatch(r'-?\d+(\.0+)?', got) and re.fullmatch(r'-?\d+(\.0+)?', want):
            return float(got) == float(want)
        return got.lower() == want.lower()

    m = _COND_ITEM_NOT_BLANK.search(s)
    if m:
        idx = colmap.get(m.group(1).lower())
        if idx is None or idx >= len(row_values):
            return None
        v = row_values[idx]
        return v is not None and str(v).strip() != ''

    m = _COND_CIC_LIST.match(s)
    if m and cic_code:
        wanted = [w.strip().upper() for w in m.group(1).split(',') if w.strip()]
        tail = cic_code.strip().upper()[2:]        # CIC is 2-letter country + 2 chars
        for w in wanted:
            if tail.startswith(w) or tail == w:
                return True
        return False

    return None


# --- field types, inferred from the template's CODIFICATION column -------------

_TYPE_RULES = [
    ('cic', re.compile(r'\bCIC\s*code\b', re.I)),
    ('currency', re.compile(r'ISO\s*4217', re.I)),
    ('country', re.compile(r'ISO\s*3166', re.I)),
    ('date', re.compile(r'YYYY\s*-?\s*MM\s*-?\s*DD|ISO\s*8601', re.I)),
    ('integer', re.compile(r'^\s*integer\b', re.I)),
    ('number', re.compile(r'^\s*num(ber)?\b|floating\s+decimal|^\s*[01]\s*=', re.I)),
    ('flag', re.compile(r'^\s*(Y\s*/\s*N|N\s*/\s*Y)|"?Y"?\s*(or|/)\s*"?N"?', re.I)),
]

# Fields whose content is an identifier with its own check digit, when the
# companion "type of code" field says so.
ISIN_TOKENS = {'14'}
LEI_TOKENS = {'47', '119'}


def field_type(f: 'Field') -> str:
    """One of: cic, currency, country, date, integer, number, flag, text."""
    text = ' '.join(str(f.codification or '').split())
    name = f.name.lower()
    if f.token == '26':
        return 'number'
    if 'date' in name and not text:
        return 'date'
    for kind, rx in _TYPE_RULES:
        if rx.search(text):
            return kind
    return 'text'
