"""The check framework: context, result and registry."""
from __future__ import annotations

from dataclasses import dataclass, field as dc_field
from typing import Any, Callable, Dict, List, Optional

CHECKS: Dict[str, 'CheckSpec'] = {}

DIMENSION_OF_PREFIX = {'S': 'structure', 'C': 'completeness', 'V': 'validity', 'K': 'consistency'}


@dataclass
class CheckSpec:
    id: str
    title: str
    dimension: str
    fn: Callable


def check(check_id: str, title: str):
    """Register a check. The prefix of the id decides its dimension."""
    def deco(fn):
        CHECKS[check_id] = CheckSpec(check_id, title, DIMENSION_OF_PREFIX[check_id[0]], fn)
        return fn
    return deco


@dataclass
class Issue:
    row: Optional[int]            # 1-based row in the sheet, None for file-level
    column: str                   # column header or field token
    found: Any
    expected: str

    def as_row(self):
        return [self.row, self.column, '' if self.found is None else str(self.found)[:120],
                self.expected]


@dataclass
class CheckResult:
    id: str
    title: str
    dimension: str
    f: float = 0.0                # 0..1 failure measure
    n_affected: int = 0
    n_applicable: int = 0
    detail: str = ''
    issues: List[Issue] = dc_field(default_factory=list)
    evaluated: bool = True

    @property
    def status(self) -> str:
        if not self.evaluated:
            return 'Not evaluated'
        if self.f <= 0:
            return 'Pass'
        return 'Fail' if self.f >= 0.5 else 'Minor'

    def add(self, row, column, found, expected):
        if len(self.issues) < 200:           # keep the issue log usable
            self.issues.append(Issue(row, column, found, expected))


def not_evaluated(spec: 'CheckSpec', why: str) -> CheckResult:
    return CheckResult(spec.id, spec.title, spec.dimension, f=0.0, detail=why, evaluated=False)


def share(affected: int, applicable: int) -> float:
    if applicable <= 0:
        return 0.0
    return max(0.0, min(1.0, affected / float(applicable)))


def band(deviation: float, tolerance: float, hard_limit: float) -> float:
    """Aggregate deviation turned into 0..1: nothing inside tolerance, 1 at the
    hard limit."""
    if deviation <= tolerance:
        return 0.0
    span = max(hard_limit - tolerance, 1e-9)
    return max(0.0, min(1.0, (deviation - tolerance) / span))


@dataclass
class Context:
    """Everything a check needs about one submission."""
    path: str
    table: Any                     # loader.Table
    cmap: Any                      # mapper.ColumnMap
    fields: Any                    # rules.FieldSet
    cfg: Any                       # config.Config
    version: str = 'unknown'
    _cic: Optional[List[Optional[str]]] = None

    # -- row helpers -------------------------------------------------------
    @property
    def rows(self):
        return self.table.rows

    @property
    def n_rows(self) -> int:
        return len(self.table.rows)

    def sheet_row(self, i: int) -> int:
        """1-based row number in the sheet for body row i."""
        return self.table.header_row + 1 + i

    def value(self, row, tokenid):
        return self.cmap.value(row, tokenid)

    def column_name(self, tokenid) -> str:
        i = self.cmap.get(tokenid)
        if i is None or i >= len(self.table.header):
            f = self.fields.get(tokenid)
            return f.name if f else str(tokenid)
        h = self.table.header[i]
        return str(h) if h is not None else str(tokenid)

    def cic_of(self, i: int) -> Optional[str]:
        """CIC category letter for body row i, cached."""
        from ..util import cic_category
        if self._cic is None:
            self._cic = [cic_category(self.value(r, '12')) for r in self.table.rows]
        return self._cic[i] if i < len(self._cic) else None

    def opt(self, check_id: str, key: str, default=None):
        return self.cfg.check(check_id).get(key, default)


def run_all(ctx: Context) -> Dict[str, CheckResult]:
    """Run every registered check in id order."""
    out: Dict[str, CheckResult] = {}
    for cid in sorted(CHECKS, key=lambda k: (k[0], int(k[1:]))):
        spec = CHECKS[cid]
        try:
            res = spec.fn(ctx)
        except Exception as exc:                      # a broken check must not stop a run
            res = CheckResult(spec.id, spec.title, spec.dimension, f=0.0, evaluated=False,
                              detail='check raised %s: %s' % (type(exc).__name__, str(exc)[:120]))
        if res is None:
            res = not_evaluated(spec, 'nothing to evaluate')
        out[cid] = res
    return out
