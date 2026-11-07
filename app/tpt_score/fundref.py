"""Fund reference: who is the GP behind each submission.

The GP is identified from the fund name in the submission's FILE NAME (for
example "Meridian PE Access Fund_TPT_2026Q3.xlsx"), not from the portfolio
name inside the TPT table. ``match_filename`` does that; ``match`` compares two
plain names and is kept for name-to-name lookups.

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
from typing import List, Tuple

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
    'ptrs': 'partners', 'hldgs': 'holdings', 'secs': 'secondaries', 'cre': 'commercial real estate',
    '1': 'i', '2': 'ii', '3': 'iii', '4': 'iv', '5': 'v', '6': 'vi', '7': 'vii', '8': 'viii',
}
_NUMERALS = {'i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x'}

CLOSE_CUTOFF = 0.82        # similarity at or above this is a close match
FILE_CUTOFF = 0.70         # file names abbreviate more, so a lower bar (weighted by rarity)
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
            return 'Matched: %s' % self.ref_fund
        if self.status == 'close':
            return 'Close match: %s (%.0f%%)' % (self.ref_fund, 100 * self.similarity)
        if self.status == 'ambiguous':
            return 'Ambiguous: %s' % self.alternatives
        if self.status == 'no name':
            return 'No fund name in the file name'
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


# Words in file names that are never part of a fund name.
_FILE_NOISE = {'tpt', 'v5', 'v6', 'v7', 'v8', 'final', 'draft', 'copy', 'report', 'template',
               'data', 'file', 'q1', 'q2', 'q3', 'q4', 'h1', 'h2', 'fy', 'xlsx', 'xlsm', 'rev',
               'updated', 'new', 'old', 'version', 'submission'}


def filename_tokens(filename: str) -> List[str]:
    """Words of a file name: extension dropped, CamelCase and separators split,
    dates and noise words removed, abbreviations expanded."""
    stem = re.sub(r'\.[A-Za-z0-9]{3,4}$', '', str(filename))
    stem = re.sub(r'([a-z])([A-Z])', r'\1 \2', stem)          # MeridianAccess -> Meridian Access
    stem = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', stem)    # PEAccess -> PE Access
    words = []
    for w in re.split(r'[^A-Za-z0-9]+', stem.lower().replace('&', ' and ')):
        if not w or w in _FILE_NOISE or len(re.findall(r'\d', w)) >= 3 or \
                re.fullmatch(r'v\d+|q\d|[a-z]{3}\d{2}|\d{1,2}[a-z]{3}', w):
            continue
        w = _ABBREV.get(w, w)
        words.extend(w.split())
    return [w for w in words if w not in _NOISE]


def _stem(w: str) -> str:
    """Crude singular form, so "secondaries" meets "secondary"."""
    if len(w) > 4 and w.endswith('ies'):
        return w[:-3] + 'y'
    if len(w) > 3 and w.endswith('s') and not w.endswith('ss'):
        return w[:-1]
    return w


def _token_hit(tok: str, words: List[str]) -> float:
    """1.0 for the word itself, the similarity for a near spelling, else 0."""
    if tok in words or _stem(tok) in {_stem(w) for w in words}:
        return 1.0
    if tok in _NUMERALS:
        return 0.0                       # fund numbers must match exactly
    best = max((difflib.SequenceMatcher(None, tok, w).ratio() for w in words), default=0.0)
    return best if best >= 0.85 else 0.0


class FundReference:
    def __init__(self, funds: List[RefFund], path: str = ''):
        import math
        self.funds = funds
        self.path = path
        # Rare words (a fund's own name) weigh more than words most funds share
        # ("infrastructure", "debt", a house prefix).
        df = {}
        for f in funds:
            for w in set(f.key.split()):
                df[w] = df.get(w, 0) + 1
        n = max(1, len(funds))
        self.weight = {w: math.log(1 + n / c) for w, c in df.items()}

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

    def match_filename(self, filename: str) -> FundMatch:
        """Find the reference fund whose name appears in ``filename``."""
        words = filename_tokens(filename)
        if not words:
            return FundMatch('no name')
        file_nums = set(words) & _NUMERALS
        scored = []
        for f in self.funds:
            ref = [w for w in f.key.split()]
            if not ref:
                continue
            # A numbered fund needs its number in the file name, and a file naming
            # a fund number cannot belong to a fund with a different one.
            ref_nums = set(ref) & _NUMERALS
            if not ref_nums <= file_nums or (file_nums and ref_nums and file_nums != ref_nums):
                continue
            hits = [_token_hit(t, words) for t in ref]
            found = sum(1 for h in hits if h)
            if found < min(2, len(ref)):
                continue
            wts = [self.weight.get(t, 1.0) for t in ref]
            sim = sum(h * w for h, w in zip(hits, wts)) / sum(wts)
            exact = all(h == 1.0 for h in hits)
            scored.append((sim, found, exact, f))
        if not scored:
            return FundMatch('not found')
        # Best coverage first; among equals, the fund that explains more words.
        scored.sort(key=lambda x: (-x[0], -x[1]))
        sim, found, exact, best = scored[0]
        if sim < FILE_CUTOFF:
            return FundMatch('not found', similarity=sim, alternatives='nearest: %s' % best.fund)
        rivals = [f for s2, n2, _, f in scored[1:]
                  if sim - s2 <= AMBIGUOUS_GAP and n2 == found and f.gp != best.gp]
        if rivals:
            return FundMatch('ambiguous', similarity=sim,
                             alternatives=' / '.join([best.fund] + [r.fund for r in rivals[:2]]))
        if exact:
            return FundMatch('exact', best.gp, best.fund, 1.0, same_text=False)
        return FundMatch('close', best.gp, best.fund, sim)


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
