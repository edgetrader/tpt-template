"""Fund reference: who is the GP behind each submission.

The reference is an Excel workbook with a "Fund Name" tab holding a table with
at least two columns, GP Name and Fund Name. The fund name a GP wrote in the
TPT (field 3, falling back to field 1) is matched to that table. GPs spell
their funds loosely ("Meridian PE Access Fund, L.P." for "Meridian Private
Equity Access Fund"), so the match is fuzzy:

* exact     the names are the same once case, punctuation and legal-form
            words (LP, SCSp, SICAV, ...) are ignored
* close     not identical but clearly the same fund; shown so a person can
            confirm it
* ambiguous two reference funds match about equally well
* not found nothing in the reference is close enough

Fund numbering is kept: "Fund II" never matches "Fund III".
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass
from typing import List, Optional, Tuple

import openpyxl

SHEET_NAME = 'fund name'

# Words that carry no identity: legal forms and filler.
_NOISE = {
    'the', 'fund', 'funds', 'lp', 'l', 'p', 'llp', 'llc', 'ltd', 'limited', 'inc', 'sa', 'sarl',
    'scsp', 'scs', 'sca', 'sicav', 'sicaf', 'raif', 'sif', 'fcp', 'eltif', 'ucits', 'aif',
    'plc', 'gmbh', 'kg', 'co', 'and', 'of', 'a', 'company', 'partnership',
}
# Common abbreviations, expanded before comparing.
_ABBREV = {
    'pe': 'private equity', 'infra': 'infrastructure', 're': 'real estate',
    'opps': 'opportunities', 'opp': 'opportunities', 'intl': 'international',
    'mgmt': 'management', 'cap': 'capital', 'sr': 'senior', 'dl': 'direct lending',
    'ptrs': 'partners', 'hldgs': 'holdings', 'secs': 'secondaries',
    '1': 'i', '2': 'ii', '3': 'iii', '4': 'iv', '5': 'v', '6': 'vi', '7': 'vii', '8': 'viii',
}
_NUMERALS = {'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'}

CLOSE_CUTOFF = 0.82        # similarity at or above this is a close match
AMBIGUOUS_GAP = 0.03       # two candidates this near each other are ambiguous


@dataclass
class RefFund:
    gp: str
    fund: str
    key: str                 # normalised fund name
    tokens: frozenset


@dataclass
class FundMatch:
    status: str              # exact, close, ambiguous, not found, no name
    gp: str = ''
    ref_fund: str = ''
    similarity: float = 0.0
    alternatives: str = ''
    same_text: bool = True     # the GP wrote the name exactly as the reference does

    @property
    def found(self) -> bool:
        return self.status in ('exact', 'close')

    @property
    def label(self) -> str:
        if self.status == 'exact':
            return 'Matched' if self.same_text else 'Matched: %s' % self.ref_fund
        if self.status == 'close':
            return 'Close match: %s (%.0f%%)' % (self.ref_fund, 100 * self.similarity)
        if self.status == 'ambiguous':
            return 'Ambiguous: %s' % self.alternatives
        if self.status == 'no name':
            return 'No fund name in the file'
        return 'NOT IN FUND REFERENCE'


def normalise_name(name: str) -> Tuple[str, frozenset]:
    s = str(name or '').lower().replace('&', ' and ')
    s = re.sub(r'[^a-z0-9]+', ' ', s)
    words = []
    for w in s.split():
        w = _ABBREV.get(w, w)
        words.extend(w.split())
    words = [w for w in words if w not in _NOISE]
    return ' '.join(words), frozenset(words)


class FundReference:
    def __init__(self, funds: List[RefFund], path: str = ''):
        self.funds = funds
        self.path = path

    def __len__(self):
        return len(self.funds)

    def match(self, name: str) -> FundMatch:
        key, tokens = normalise_name(name)
        if not key:
            return FundMatch('no name')
        for f in self.funds:
            if f.key == key:
                same = ' '.join(str(name).lower().split()) == ' '.join(f.fund.lower().split())
                return FundMatch('exact', f.gp, f.fund, 1.0, same_text=same)
        scored = []
        for f in self.funds:
            # Numbering must agree: Fund II is a different fund from Fund III.
            if (tokens & _NUMERALS) != (f.tokens & _NUMERALS):
                continue
            seq = difflib.SequenceMatcher(None, key, f.key).ratio()
            union = tokens | f.tokens
            jac = len(tokens & f.tokens) / len(union) if union else 0.0
            contain = len(tokens & f.tokens) / min(len(tokens), len(f.tokens)) if tokens and f.tokens else 0
            sim = max(seq, (jac + contain) / 2)
            scored.append((sim, f))
        if not scored:
            return FundMatch('not found')
        scored.sort(key=lambda x: -x[0])
        best_sim, best = scored[0]
        if best_sim < CLOSE_CUTOFF:
            return FundMatch('not found', similarity=best_sim,
                             alternatives='nearest: %s' % best.fund)
        rivals = [f for s, f in scored[1:] if best_sim - s <= AMBIGUOUS_GAP and f.gp != best.gp]
        if rivals:
            return FundMatch('ambiguous', similarity=best_sim,
                             alternatives=' / '.join([best.fund] + [r.fund for r in rivals[:2]]))
        return FundMatch('close', best.gp, best.fund, best_sim)


class FundReferenceError(Exception):
    pass


def _norm_head(v) -> str:
    return re.sub(r'[^a-z]', '', str(v or '').lower())


def load_fund_reference(path: str) -> FundReference:
    """Read the table on the "Fund Name" tab (name matched loosely)."""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        raise FundReferenceError('cannot open fund reference %s: %s' % (path, exc))
    try:
        names = {re.sub(r'\s+', ' ', n.strip().lower()): n for n in wb.sheetnames}
        sheet = names.get(SHEET_NAME) or next(
            (n for k, n in names.items() if k.replace(' ', '') in ('fundname', 'fundnames')), None)
        if sheet is None:
            raise FundReferenceError('fund reference %s has no "Fund Name" tab (tabs: %s)'
                                     % (path, ', '.join(wb.sheetnames)))
        rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    finally:
        wb.close()

    gp_col = fund_col = None
    start = 0
    for i, row in enumerate(rows[:30]):
        heads = [_norm_head(v) for v in row]
        if 'gpname' in heads and 'fundname' in heads:
            gp_col, fund_col, start = heads.index('gpname'), heads.index('fundname'), i + 1
            break
    if gp_col is None:
        raise FundReferenceError('the "Fund Name" tab in %s needs a header row with '
                                 '"GP Name" and "Fund Name"' % path)

    funds, seen = [], set()
    for row in rows[start:]:
        gp = str(row[gp_col]).strip() if gp_col < len(row) and row[gp_col] is not None else ''
        fund = str(row[fund_col]).strip() if fund_col < len(row) and row[fund_col] is not None else ''
        if not fund:
            continue
        key, tokens = normalise_name(fund)
        if not key or (key, gp) in seen:
            continue
        seen.add((key, gp))
        funds.append(RefFund(gp, fund, key, tokens))
    if not funds:
        raise FundReferenceError('the "Fund Name" tab in %s lists no funds' % path)
    return FundReference(funds, path)


def tpt_fund_name(fs) -> str:
    """The fund name the GP wrote, falling back to the fund identifier."""
    for v in (fs.fund_name, fs.fund_id):
        if v and str(v).strip():
            return str(v).strip()
    return ''
