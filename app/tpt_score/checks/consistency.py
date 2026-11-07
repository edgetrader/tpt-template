"""K1 to K6: do the values agree with each other and with the fund totals."""
from __future__ import annotations

import collections

from ..util import as_date, as_number, text_of
from .base import CheckResult, Context, band, check, share
from .completeness import FUND_LEVEL


@check('K1', 'Fund-level fields identical on every row')
def k1(ctx: Context) -> CheckResult:
    res = CheckResult('K1', 'Fund-level fields identical on every row', 'consistency')
    tokens = [t for t in FUND_LEVEL if ctx.cmap.has(t)]
    if not tokens or ctx.n_rows < 2:
        res.detail = 'fewer than two positions, or no fund-level columns'
        res.evaluated = bool(tokens) and ctx.n_rows >= 2
        return res
    # The majority value of each fund-level column is taken as the intended one.
    majority = {}
    for t in tokens:
        counts = collections.Counter(text_of(ctx.value(r, t)) for r in ctx.rows)
        majority[t] = counts.most_common(1)[0][0]
    affected = 0
    for i, row in enumerate(ctx.rows):
        odd = [t for t in tokens if text_of(ctx.value(row, t)) != majority[t]]
        if odd:
            affected += 1
            for t in odd[:4]:
                res.add(ctx.sheet_row(i), ctx.column_name(t), ctx.value(row, t),
                        'the same value as the rest of the file (%r)' % majority[t][:40])
    res.n_applicable, res.n_affected = ctx.n_rows, affected
    res.f = share(affected, ctx.n_rows)
    res.detail = '%d of %d rows carry a fund-level value that differs from the rest' % (
        affected, ctx.n_rows)
    return res


@check('K2', 'Valuation weights sum to 100%')
def k2(ctx: Context) -> CheckResult:
    res = CheckResult('K2', 'Valuation weights sum to 100%', 'consistency')
    if not ctx.cmap.has('26') or not ctx.rows:
        res.detail = 'no valuation weight column'
        res.evaluated = False
        return res
    total, seen = 0.0, 0
    for row in ctx.rows:
        n, _ = as_number(ctx.value(row, '26'))
        if n is not None:
            total += n
            seen += 1
    if seen == 0:
        res.detail = 'no weights to add up'
        res.evaluated = False
        return res
    # Weights may arrive as fractions (1 = 100%) or as percentages.
    pct = total * 100.0 if total <= 1.5 else total
    deviation = abs(pct - 100.0)
    tol = float(ctx.opt('K2', 'tolerance_pp', 0.5))
    hard = float(ctx.opt('K2', 'hard_limit_pp', 5.0))
    res.f = band(deviation, tol, hard)
    res.n_applicable, res.n_affected = seen, (1 if res.f > 0 else 0)
    res.detail = 'weights add up to %.3f%% (%.3f percentage points off, tolerance %.2f)' % (
        pct, deviation, tol)
    if res.f > 0:
        res.add(None, ctx.column_name('26'), '%.4f%%' % pct, '100%% within %.2f points' % tol)
    return res


@check('K3', 'Market values add up to the fund NAV')
def k3(ctx: Context) -> CheckResult:
    res = CheckResult('K3', 'Market values add up to the fund NAV', 'consistency')
    if not (ctx.cmap.has('24') and ctx.cmap.has('5')) or not ctx.rows:
        res.detail = 'market value in portfolio currency or NAV column missing'
        res.evaluated = False
        return res
    nav = None
    for row in ctx.rows:
        n, _ = as_number(ctx.value(row, '5'))
        if n is not None:
            nav = n
            break
    total, seen = 0.0, 0
    for row in ctx.rows:
        n, _ = as_number(ctx.value(row, '24'))
        if n is not None:
            total += n
            seen += 1
    if nav is None or seen == 0:
        res.detail = 'no readable NAV, or no market values to add up'
        res.evaluated = False
        return res
    if nav == 0:
        # A fund reporting a NAV of zero while holding positions is wrong on its
        # face, and there is no denominator to measure a deviation against.
        res.n_applicable = seen
        if total != 0:
            res.f, res.n_affected = 1.0, 1
            res.detail = ('NAV is reported as zero while the positions add up to %.2f'
                          % total)
            res.add(None, ctx.column_name('5'), 0, 'the fund NAV, which cannot be zero '
                                                   'while positions are held')
        else:
            res.detail = 'NAV and market values are both zero'
        return res
    deviation = abs(total - nav) / abs(nav) * 100.0
    tol = float(ctx.opt('K3', 'tolerance_pct', 0.5))
    hard = float(ctx.opt('K3', 'hard_limit_pct', 5.0))
    res.f = band(deviation, tol, hard)
    res.n_applicable, res.n_affected = seen, (1 if res.f > 0 else 0)
    res.detail = 'market values add up to %.2f against a NAV of %.2f (%.3f%% apart, tolerance %.2f%%)' % (
        total, nav, deviation, tol)
    if res.f > 0:
        res.add(None, ctx.column_name('24'), '%.2f' % total,
                'close to the NAV of %.2f' % nav)
    return res


@check('K4', 'Quotation and portfolio market values agree through one FX rate')
def k4(ctx: Context) -> CheckResult:
    res = CheckResult('K4', 'Quotation and portfolio market values agree through one FX rate',
                      'consistency')
    need = ('22', '24', '21')
    if not all(ctx.cmap.has(t) for t in need) or not ctx.rows:
        res.detail = 'market value or quotation currency column missing'
        res.evaluated = False
        return res
    # One implied rate per quotation currency, taken as the median of the rows.
    rates = collections.defaultdict(list)
    per_row = []
    for i, row in enumerate(ctx.rows):
        qc, _ = as_number(ctx.value(row, '22'))
        pc, _ = as_number(ctx.value(row, '24'))
        ccy = text_of(ctx.value(row, '21')).upper()
        if qc in (None, 0) or pc is None or not ccy:
            per_row.append(None)
            continue
        r = pc / qc
        rates[ccy].append(r)
        per_row.append((i, ccy, r))
    median = {}
    for ccy, rs in rates.items():
        rs = sorted(rs)
        median[ccy] = rs[len(rs) // 2]
    tol = float(ctx.opt('K4', 'tolerance_pct', 1.0)) / 100.0
    applicable = affected = 0
    for item in per_row:
        if item is None:
            continue
        i, ccy, r = item
        ref = median.get(ccy)
        if not ref:
            continue
        applicable += 1
        if abs(r - ref) > abs(ref) * tol:
            affected += 1
            res.add(ctx.sheet_row(i), ctx.column_name('24'), 'implied rate %.6f' % r,
                    'about %.6f, the rate the rest of the file uses for %s' % (ref, ccy))
    res.n_applicable, res.n_affected = applicable, affected
    res.f = share(affected, applicable)
    if applicable == 0:
        res.evaluated = False
        res.detail = 'no row allows an implied rate to be worked out'
    else:
        res.detail = '%d of %d rows imply an FX rate away from the rest of the file' % (
            affected, applicable)
    return res


@check('K5', 'No duplicated positions')
def k5(ctx: Context) -> CheckResult:
    res = CheckResult('K5', 'No duplicated positions', 'consistency')
    if ctx.n_rows < 2:
        res.detail = 'fewer than two positions'
        res.evaluated = False
        return res
    # One pass per index: building the key the second time for every duplicate
    # found turns this check quadratic on files that repeat an identifier often.
    key_token = '14' if ctx.cmap.has('14') else None
    by_row = collections.defaultdict(list)
    by_key = collections.defaultdict(list)
    for i, row in enumerate(ctx.rows):
        by_row[tuple(text_of(v) for v in row)].append(i)
        if key_token:
            k = text_of(ctx.value(row, key_token))
            if k:
                by_key[k].append(i)

    affected = set()
    for groups in (by_row, by_key):
        for idxs in groups.values():
            if len(idxs) > 1:
                affected.update(idxs)
    for i in sorted(affected)[:40]:
        res.add(ctx.sheet_row(i), ctx.column_name(key_token) if key_token else 'row',
                text_of(ctx.value(ctx.rows[i], key_token)) if key_token else '(identical row)',
                'one row per position')
    res.n_applicable, res.n_affected = ctx.n_rows, len(affected)
    res.f = share(len(affected), ctx.n_rows)
    res.detail = '%d of %d positions are duplicated' % (len(affected), ctx.n_rows)
    return res


@check('K6', 'Date logic: valuation date not after reporting date, and not stale')
def k6(ctx: Context) -> CheckResult:
    res = CheckResult('K6', 'Date logic: valuation date not after reporting date, and not stale',
                      'consistency')
    if not (ctx.cmap.has('6') and ctx.cmap.has('7')) or not ctx.rows:
        res.detail = 'valuation or reporting date column missing'
        res.evaluated = False
        return res
    val = rep = None
    for row in ctx.rows:
        d, _ = as_date(ctx.value(row, '6'))
        if d and val is None:
            val = d
        d2, _ = as_date(ctx.value(row, '7'))
        if d2 and rep is None:
            rep = d2
        if val and rep:
            break
    if not val or not rep:
        res.detail = 'no readable valuation or reporting date'
        res.evaluated = False
        return res
    stale_days = int(ctx.opt('K6', 'stale_days', 100))
    problems = []
    if val > rep:
        problems.append('valuation date %s is after the reporting date %s' % (val, rep))
    gap = (rep - val).days
    if gap > stale_days:
        problems.append('valuation date is %d days before the reporting date (limit %d)' % (gap, stale_days))
    res.n_applicable = 1
    if problems:
        res.f = 1.0
        res.n_affected = 1
        res.detail = '; '.join(problems)
        res.add(None, ctx.column_name('6'), str(val),
                'on or shortly before the reporting date %s' % rep)
    else:
        res.detail = 'valuation %s, reporting %s' % (val, rep)
    return res
