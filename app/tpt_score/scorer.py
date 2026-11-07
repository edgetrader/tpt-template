"""Turn check results into a 0-100 score, a rating and a reason."""
from __future__ import annotations

import collections
import datetime as dt
import os
from dataclasses import dataclass, field as dc_field
from typing import Any, Dict, List, Optional

from .checks import Context, run_all
from .config import Config
from .loader import UnreadableFile, load_table
from .mapper import detect_version, map_columns
from .profile import profile_columns
from .remarks import FINDING_LABELS, column_findings

FINDING_KEYS = [k for k, _ in FINDING_LABELS]

DIMENSIONS = ['structure', 'completeness', 'validity', 'consistency']
DIMENSION_LABEL = {'structure': 'Structure', 'completeness': 'Completeness',
                   'validity': 'Validity and format', 'consistency': 'Consistency'}


@dataclass
class CheckScore:
    id: str
    title: str
    dimension: str
    f: float
    raw_f: float
    weight: float
    critical: bool
    status: str
    detail: str
    n_affected: int
    n_applicable: int
    points_lost: float          # on the 0-100 total scale
    issues: List[Any] = dc_field(default_factory=list)


@dataclass
class DimensionScore:
    name: str
    weight: float
    score: float
    contribution: float

    @property
    def label(self) -> str:
        return DIMENSION_LABEL[self.name]


@dataclass
class FileScore:
    path: str
    filename: str
    score: float = 0.0
    rating: str = 'Reject'
    gates: List[str] = dc_field(default_factory=list)
    cap: Optional[float] = None
    dimensions: Dict[str, DimensionScore] = dc_field(default_factory=dict)
    checks: Dict[str, CheckScore] = dc_field(default_factory=dict)
    version: str = 'unknown'
    sheet: str = ''
    positions: int = 0
    mapped_columns: int = 0
    header_cell: str = ''
    recovery: List[str] = dc_field(default_factory=list)
    error: str = ''
    fund_name: str = ''
    fund_id: str = ''
    rules_version: str = ''
    scored_at: str = ''
    missing_columns: List[str] = dc_field(default_factory=list)       # mandatory fields absent
    renamed_columns: List[Any] = dc_field(default_factory=list)       # (header found, template name)
    duplicate_columns: List[str] = dc_field(default_factory=list)
    extra_columns: List[str] = dc_field(default_factory=list)
    missing_optional: List[str] = dc_field(default_factory=list)      # non-mandatory fields absent
    blank_mandatory: List[Any] = dc_field(default_factory=list)       # (column, empty cells, applicable cells)
    extra_tables: List[str] = dc_field(default_factory=list)          # other TPT tables found
    type_mismatches: List[Any] = dc_field(default_factory=list)
    date_formats: List[Any] = dc_field(default_factory=list)

    @property
    def remarks(self) -> str:
        from .remarks import build_remarks
        return build_remarks(self)

    @property
    def readable(self) -> bool:
        return not self.error

    @property
    def top_fixes(self) -> List[CheckScore]:
        bad = [c for c in self.checks.values() if c.points_lost > 0]
        return sorted(bad, key=lambda c: -c.points_lost)[:5]

    def as_summary_row(self) -> List[Any]:
        return [self.filename, round(self.score, 1), self.rating, self.version, self.positions,
                self.mapped_columns,
                round(self.dimensions['structure'].score, 1) if self.dimensions else '',
                round(self.dimensions['completeness'].score, 1) if self.dimensions else '',
                round(self.dimensions['validity'].score, 1) if self.dimensions else '',
                round(self.dimensions['consistency'].score, 1) if self.dimensions else '',
                '; '.join(self.gates), self.error, self.sheet,
                *[' | '.join(column_findings(self)[k]) for k in FINDING_KEYS],
                self.remarks]


def _column_findings(cmap, fields) -> Dict[str, Any]:
    """Which columns are missing, renamed, repeated or extra, for the remarks."""
    missing, optional = [], []
    for f in fields.fields:
        if not cmap.has(f.token):
            (missing if f.mandatory else optional).append(f.name)
    renamed = []
    for c in cmap.renamed_columns:
        f = fields.get(c.token)
        renamed.append((str(c.header), f.name if f else c.token))
    return dict(missing_columns=missing, missing_optional=optional, renamed_columns=renamed,
                duplicate_columns=[str(c.header) for c in cmap.duplicate_columns],
                extra_columns=[str(c.header).strip() or '(blank header)'
                               for c in cmap.unknown_columns])


def _col_letter(n: int) -> str:
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _zero(path: str, cfg: Config, reason: str, gate: str, **kw) -> FileScore:
    fs = FileScore(path=path, filename=os.path.basename(path), score=0.0, rating='Reject',
                   gates=['%s: %s' % (gate, reason)], error=reason,
                   rules_version=cfg.rules_version,
                   scored_at=dt.datetime.now().isoformat(timespec='seconds'))
    for name in DIMENSIONS:
        w = float(cfg.dimensions[name])
        fs.dimensions[name] = DimensionScore(name, w, 0.0, 0.0)
    for k, v in kw.items():
        setattr(fs, k, v)
    return fs


def score_file(path: str, fields, cfg: Config) -> FileScore:
    """Load, check and score one submission."""
    now = dt.datetime.now().isoformat(timespec='seconds')

    # Gate G1: the file cannot be opened at all.
    try:
        table = load_table(path, fields)
    except UnreadableFile as exc:
        return _zero(path, cfg, str(exc), 'G1')

    cmap = map_columns(table.header, fields)
    version = detect_version(table.header, table.rows, cmap, fields)

    fs_kw = _column_findings(cmap, fields)
    fs_kw.update(sheet=table.sheet, positions=len(table.rows),
                 mapped_columns=cmap.n_mapped, version=version,
                 recovery=list(table.recovery), extra_tables=list(table.extra_tables),
                 header_cell='%s%d' % (_col_letter(table.first_col), table.header_row))

    # Gate G2: no header, or too few recognisable columns to judge anything.
    min_cols = int(cfg.gates.get('min_mapped_columns', 10))
    if not table.header or cmap.n_mapped < min_cols:
        return _zero(path, cfg,
                     'only %d template columns recognised (at least %d needed)' % (cmap.n_mapped, min_cols),
                     'G2', **fs_kw)

    # Gate G3: nothing to check.
    if not table.rows:
        return _zero(path, cfg, 'no position rows below the header', 'G3', **fs_kw)
    if not cmap.has('12'):
        return _zero(path, cfg, 'the CIC column (field 12) was not recognised', 'G3', **fs_kw)

    ctx = Context(path=path, table=table, cmap=cmap, fields=fields, cfg=cfg, version=version)
    results = run_all(ctx)

    fs = FileScore(path=path, filename=os.path.basename(path), rules_version=cfg.rules_version,
                   scored_at=now, **fs_kw)
    c1 = results.get('C1')
    if c1 is not None:
        for tok, (empty, total) in c1.by_column.items():
            if empty:
                f = fields.get(tok)
                fs.blank_mandatory.append((f.name if f else tok, empty, total))
        fs.blank_mandatory.sort(key=lambda x: -x[1])
    prof = profile_columns(ctx)
    fs.type_mismatches = prof['type_mismatches']
    fs.date_formats = prof['date_formats']
    fs.fund_id = _first_value(ctx, '1')
    fs.fund_name = _first_value(ctx, '3')

    fs.checks, fs.dimensions, fs.score = combine(results, cfg)

    # --- gates that cap rather than zero ------------------------------------
    caps = []
    if version != 'V7':
        cap = float(cfg.gates.get('wrong_version_cap', 40))
        caps.append((cap, 'G4: template is %s, not V7' % version))
    if fs.extra_tables:
        caps.append((float(cfg.gates.get('multiple_tables_cap', 40)),
                     'G7: %d TPT tables found, only one is allowed' % (1 + len(fs.extra_tables))))
    mm = cfg.gates.get('missing_mandatory', {})
    s3 = results.get('S3')
    if s3 is not None and s3.f > float(mm.get('share', 0.20)):
        caps.append((float(mm.get('cap', 49)),
                     'G5: %.0f%% of mandatory columns missing' % (s3.f * 100)))
    if any(c.critical and c.raw_f > 0 for c in fs.checks.values()):
        failed = sorted(c.id for c in fs.checks.values() if c.critical and c.raw_f > 0)
        caps.append((float(cfg.gates.get('critical_failure_cap', 94)),
                     'G6: critical check failed (%s)' % ', '.join(failed)))

    for cap, reason in caps:
        fs.gates.append(reason)
        if fs.score > cap:
            fs.score = cap
            fs.cap = cap
    fs.rating = cfg.rating(fs.score)
    return fs


def combine(results, cfg: Config):
    """Weights, critical floors and dimension roll-up.

    ``results`` maps check id -> CheckResult. Returns
    (checks, dimensions, total score) without applying the gates, so the
    arithmetic can be tested on its own.
    """
    floor = cfg.critical_floor
    lost_in_dim = {d: 0.0 for d in DIMENSIONS}
    checks: Dict[str, CheckScore] = {}
    for cid, res in results.items():
        w = cfg.weight(cid)
        critical = cfg.is_critical(cid)
        f = res.f
        if critical and f > 0:
            f = max(f, floor)
        lost = w * f
        lost_in_dim[res.dimension] += lost
        dim_weight = float(cfg.dimensions[res.dimension])
        checks[cid] = CheckScore(
            id=cid, title=res.title, dimension=res.dimension, f=f, raw_f=res.f, weight=w,
            critical=critical, status=res.status, detail=res.detail,
            n_affected=res.n_affected, n_applicable=res.n_applicable,
            points_lost=lost * dim_weight / 100.0, issues=list(res.issues))

    dimensions: Dict[str, DimensionScore] = {}
    total = 0.0
    for name in DIMENSIONS:
        w = float(cfg.dimensions[name])
        score = max(0.0, 100.0 - lost_in_dim[name])
        contribution = w * score / 100.0
        dimensions[name] = DimensionScore(name, w, score, contribution)
        total += contribution
    return checks, dimensions, max(0.0, min(100.0, total))


def _first_value(ctx: Context, token: str) -> str:
    """The value a fund-level field holds on most rows.

    Taking the first row was fooled by a stray flag row (``M``) or a repeated
    label row left in the body, so the most common value wins instead, and
    values that are a template label or a single code letter are skipped.
    """
    from .rules import normalise
    labels = set(ctx.fields.by_name)
    seen = collections.Counter()
    for row in ctx.rows:
        v = ctx.value(row, token)
        t = '' if v is None else str(v).strip()
        if len(t) < 3 or normalise(t) in labels or t.lower() in ('text', 'n/a', 'none', 'null'):
            continue
        if t.lower().startswith(('name of the', 'portfolio identifying')):
            continue
        seen[t] += 1
    return seen.most_common(1)[0][0] if seen else ''
