"""C1 to C5: are the required values there."""
from __future__ import annotations

from ..rules import evaluate_condition
from ..util import is_blank, is_placeholder
from .base import CheckResult, Context, check, share

# Fields whose value describes the fund, not the position, so they must repeat
# unchanged on every row. 1000 is the template version label.
FUND_LEVEL = ['1', '2', '3', '4', '5', '6', '7', '8', '8b', '9', '10', '11', '1000']


def _applicable(ctx: Context, token: str, row, cic: str):
    """True / False / None (cannot be decided) for a conditional field."""
    f = ctx.fields.get(token)
    if f is None:
        return False
    verdict = f.applies_to_cic(cic)
    if verdict is True:
        return True
    if verdict is False:
        return False
    if isinstance(verdict, str):
        return evaluate_condition(verdict, row, ctx.cmap.index, ctx.value(row, '12'))
    return None


@check('C1', 'Mandatory fields populated')
def c1(ctx: Context) -> CheckResult:
    res = CheckResult('C1', 'Mandatory fields populated', 'completeness')
    tokens = [t for t in ctx.fields.mandatory_tokens if ctx.cmap.has(t)]
    if not tokens or not ctx.rows:
        res.detail = 'no mandatory columns mapped'
        res.evaluated = bool(tokens)
        return res
    applicable = affected = 0
    for i, row in enumerate(ctx.rows):
        cic = ctx.cic_of(i)
        for t in tokens:
            verdict = _applicable(ctx, t, row, cic)
            if verdict is False:
                continue                      # not applicable to this CIC category
            applicable += 1
            col = res.by_column.setdefault(t, [0, 0])
            col[1] += 1
            if is_blank(ctx.value(row, t)):
                affected += 1
                col[0] += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), '(empty)', 'a value')
    res.n_applicable, res.n_affected = applicable, affected
    res.f = share(affected, applicable)
    res.detail = '%d of %d applicable mandatory cells are empty' % (affected, applicable)
    return res


@check('C2', 'Conditional fields populated when their trigger applies')
def c2(ctx: Context) -> CheckResult:
    res = CheckResult('C2', 'Conditional fields populated when their trigger applies', 'completeness')
    tokens = [t for t in ctx.fields.conditional_tokens if ctx.cmap.has(t)]
    if not tokens or not ctx.rows:
        res.detail = 'no conditional columns mapped'
        res.evaluated = bool(tokens)
        return res
    applicable = affected = undecided = 0
    for i, row in enumerate(ctx.rows):
        cic = ctx.cic_of(i)
        for t in tokens:
            verdict = _applicable(ctx, t, row, cic)
            if verdict is None:
                undecided += 1
                continue                      # free-text condition: no penalty
            if verdict is False:
                continue
            applicable += 1
            if is_blank(ctx.value(row, t)):
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), '(empty)',
                        'a value: the condition for this field applies')
    res.n_applicable, res.n_affected = applicable, affected
    res.f = share(affected, applicable)
    if applicable == 0:
        res.evaluated = undecided == 0
        res.detail = 'no conditional field could be decided (%d free-text conditions)' % undecided
    else:
        res.detail = '%d of %d triggered conditional cells are empty (%d conditions not evaluated)' % (
            affected, applicable, undecided)
    return res


@check('C3', 'Instrument identifier and CIC on every position')
def c3(ctx: Context) -> CheckResult:
    res = CheckResult('C3', 'Instrument identifier and CIC on every position', 'completeness')
    tokens = [t for t in ctx.opt('C3', 'fields', ['12', '14']) if ctx.cmap.has(t)]
    if not tokens or not ctx.rows:
        res.detail = 'identifier columns not mapped'
        res.evaluated = False
        return res
    affected = 0
    for i, row in enumerate(ctx.rows):
        missing = [t for t in tokens if is_blank(ctx.value(row, t))]
        if missing:
            affected += 1
            for t in missing:
                res.add(ctx.sheet_row(i), ctx.column_name(t), '(empty)', 'a value on every position')
    res.n_applicable, res.n_affected = ctx.n_rows, affected
    res.f = share(affected, ctx.n_rows)
    res.detail = '%d of %d positions lack the CIC code or the instrument identifier' % (affected, ctx.n_rows)
    return res


@check('C4', 'Fund-level fields on every row')
def c4(ctx: Context) -> CheckResult:
    res = CheckResult('C4', 'Fund-level fields on every row', 'completeness')
    tokens = [t for t in FUND_LEVEL if ctx.cmap.has(t) and ctx.fields.get(t) and
              ctx.fields.get(t).flag in ('M', 'C')]
    if not tokens or not ctx.rows:
        res.detail = 'no fund-level columns mapped'
        res.evaluated = bool(tokens)
        return res
    affected = 0
    for i, row in enumerate(ctx.rows):
        missing = [t for t in tokens if is_blank(ctx.value(row, t))]
        if missing:
            affected += 1
            for t in missing[:4]:
                res.add(ctx.sheet_row(i), ctx.column_name(t), '(empty)',
                        'fund-level values repeated on every row')
    res.n_applicable, res.n_affected = ctx.n_rows, affected
    res.f = share(affected, ctx.n_rows)
    res.detail = '%d of %d rows are missing a fund-level value' % (affected, ctx.n_rows)
    return res


@check('C5', 'No placeholders used as values')
def c5(ctx: Context) -> CheckResult:
    res = CheckResult('C5', 'No placeholders used as values', 'completeness')
    extra = ctx.opt('C5', 'placeholders', None)
    # A placeholder is a placeholder wherever it sits, so every mapped template
    # column counts, not only the mandatory ones.
    tokens = list(ctx.cmap.mapped_tokens)
    if not tokens or not ctx.rows:
        res.detail = 'no template columns mapped'
        res.evaluated = bool(tokens)
        return res
    populated = affected = 0
    for i, row in enumerate(ctx.rows):
        for t in tokens:
            v = ctx.value(row, t)
            if is_blank(v):
                continue
            populated += 1
            if is_placeholder(v, extra):
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, 'a real value, not a placeholder')
    res.n_applicable, res.n_affected = populated, affected
    res.f = share(affected, populated)
    res.detail = '%d of %d populated cells hold a placeholder instead of a value' % (affected, populated)
    return res
