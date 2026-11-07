"""V1 to V6: do the values have the right type, format and code."""
from __future__ import annotations

import collections

from ..rules import field_type
from ..util import (ISO_COUNTRIES, ISO_CURRENCIES, as_date, as_number, has_text_noise,
                    is_blank, is_excel_error, is_placeholder, text_of, valid_cic,
                    valid_isin, valid_lei)
from .base import CheckResult, Context, check, share

NUMERIC_TYPES = {'number', 'integer'}
FLAG_TOKENS = {'11'}                      # Y/N fields not spelled out in CODIFICATION
WEIGHT_TOKENS = {'26', '9'}               # 1 = 100%, so a value above 1.5 is suspect


def _typed_tokens(ctx: Context, kinds):
    out = []
    for t in ctx.cmap.mapped_tokens:
        f = ctx.fields.get(t)
        if f is not None and field_type(f) in kinds:
            out.append(t)
    return out


@check('V1', 'Numbers stored as numbers')
def v1(ctx: Context) -> CheckResult:
    res = CheckResult('V1', 'Numbers stored as numbers', 'validity')
    tokens = _typed_tokens(ctx, NUMERIC_TYPES)
    if not tokens or not ctx.rows:
        res.detail = 'no numeric columns mapped'
        res.evaluated = bool(tokens)
        return res
    populated = affected = 0
    for i, row in enumerate(ctx.rows):
        for t in tokens:
            v = ctx.value(row, t)
            if is_blank(v):
                continue
            populated += 1
            if is_excel_error(v):
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, 'a number, not an Excel error')
                continue
            if is_placeholder(v):
                continue                        # counted by C5, not here
            num, from_text = as_number(v)
            if num is None:
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, 'a number')
            elif from_text:
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, 'a number, not text')
    res.n_applicable, res.n_affected = populated, affected
    res.f = share(affected, populated)
    res.detail = '%d of %d populated numeric cells are text or an error value' % (affected, populated)
    return res


@check('V2', 'Dates valid and in one format')
def v2(ctx: Context) -> CheckResult:
    res = CheckResult('V2', 'Dates valid and in one format', 'validity')
    tokens = _typed_tokens(ctx, {'date'})
    if not tokens or not ctx.rows:
        res.detail = 'no date columns mapped'
        res.evaluated = bool(tokens)
        return res
    populated = affected = 0
    seen = collections.Counter()
    for i, row in enumerate(ctx.rows):
        for t in tokens:
            v = ctx.value(row, t)
            if is_blank(v):
                continue
            populated += 1
            d, kind = as_date(v)
            if d is None:
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, 'a date (YYYY-MM-DD)')
            else:
                seen[kind] += 1
    # A file that mixes formats is penalised for the minority forms as well.
    mixed = 0
    if len(seen) > 1:
        main = seen.most_common(1)[0][1]
        mixed = sum(seen.values()) - main
        res.add(None, 'date columns', ', '.join('%s x%d' % (k, n) for k, n in seen.most_common()),
                'one date format throughout')
    res.n_applicable = populated
    res.n_affected = affected + mixed
    res.f = share(affected + mixed, populated)
    res.detail = '%d invalid dates and %d in a minority format, of %d date cells' % (
        affected, mixed, populated)
    return res


@check('V3', 'Valid codes (CIC, currency, country, Y/N)')
def v3(ctx: Context) -> CheckResult:
    res = CheckResult('V3', 'Valid codes (CIC, currency, country, Y/N)', 'validity')
    kinds = {}
    for t in ctx.cmap.mapped_tokens:
        f = ctx.fields.get(t)
        if f is None:
            continue
        ft = field_type(f)
        if ft in ('cic', 'currency', 'country') or t in FLAG_TOKENS:
            kinds[t] = 'flag' if t in FLAG_TOKENS else ft
    if not kinds or not ctx.rows:
        res.detail = 'no closed-list columns mapped'
        res.evaluated = bool(kinds)
        return res
    populated = affected = 0
    for i, row in enumerate(ctx.rows):
        for t, kind in kinds.items():
            v = ctx.value(row, t)
            if is_blank(v) or is_placeholder(v):
                continue
            populated += 1
            s = text_of(v).upper()
            ok, want = True, ''
            if kind == 'cic':
                ok, want = valid_cic(v), 'a 4-character CIC code'
            elif kind == 'currency':
                ok, want = s in ISO_CURRENCIES, 'an ISO 4217 currency code'
            elif kind == 'country':
                ok, want = s in ISO_COUNTRIES, 'an ISO 3166-1 alpha-2 country code'
            elif kind == 'flag':
                ok, want = s in ('Y', 'N'), 'Y or N'
            if not ok:
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, want)
    res.n_applicable, res.n_affected = populated, affected
    res.f = share(affected, populated)
    res.detail = '%d of %d closed-list cells hold a value outside the list' % (affected, populated)
    return res


@check('V4', 'Valid ISIN and LEI codes')
def v4(ctx: Context) -> CheckResult:
    res = CheckResult('V4', 'Valid ISIN and LEI codes', 'validity')
    checks = []
    if ctx.cmap.has('14'):
        checks.append(('14', '15', 'ISIN'))         # 15 says which code type 14 holds
    for t in ('47', '119'):
        if ctx.cmap.has(t):
            checks.append((t, None, 'LEI'))
    if not checks or not ctx.rows:
        res.detail = 'no identifier columns mapped'
        res.evaluated = bool(checks)
        return res
    # The denominator is every populated identifier cell, so that one bad code
    # among many good ones costs a proportionate number of points. Only values
    # that claim a standard format are put through its check digit; a
    # proprietary code sitting in the same column is left alone.
    populated = affected = tested = 0
    for i, row in enumerate(ctx.rows):
        for token, type_token, kind in checks:
            v = ctx.value(row, token)
            if is_blank(v) or is_placeholder(v):
                continue
            populated += 1
            s = text_of(v).replace(' ', '')
            if kind == 'ISIN':
                declared = text_of(ctx.value(row, type_token)).upper() if type_token else ''
                claims_isin = declared in ('1', '1.0', 'ISIN') or \
                    (len(s) == 12 and s[:2].isalpha() and s[2:].isalnum())
                if not claims_isin:
                    continue
                tested += 1
                if not valid_isin(s):
                    affected += 1
                    res.add(ctx.sheet_row(i), ctx.column_name(token), v, 'a valid ISIN')
            else:
                if len(s) != 20:
                    continue                 # a proprietary issuer code, not an LEI
                tested += 1
                if not valid_lei(s):
                    affected += 1
                    res.add(ctx.sheet_row(i), ctx.column_name(token), v, 'a valid LEI')
    res.n_applicable, res.n_affected = populated, affected
    res.f = share(affected, populated)
    if tested == 0:
        res.evaluated = False
        res.detail = 'no value claims the ISIN or LEI format'
    else:
        res.detail = '%d of %d identifier values fail their check digit (%d tested)' % (
            affected, populated, tested)
    return res


@check('V5', 'Numbers in a sensible range')
def v5(ctx: Context) -> CheckResult:
    res = CheckResult('V5', 'Numbers in a sensible range', 'validity')
    tokens = _typed_tokens(ctx, NUMERIC_TYPES)
    if not tokens or not ctx.rows:
        res.detail = 'no numeric columns mapped'
        res.evaluated = bool(tokens)
        return res
    populated = affected = 0
    for i, row in enumerate(ctx.rows):
        for t in tokens:
            v = ctx.value(row, t)
            if is_blank(v) or is_excel_error(v) or is_placeholder(v):
                continue
            num, _ = as_number(v)
            if num is None:
                continue                         # a type problem, charged by V1
            populated += 1
            if t in WEIGHT_TOKENS and abs(num) > 1.5:
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v,
                        'a fraction where 1 = 100%, not a 0-100 percentage')
            elif t in ('5', '22', '24') and num < 0:
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, 'a value of zero or more')
    res.n_applicable, res.n_affected = populated, affected
    res.f = share(affected, populated)
    res.detail = '%d of %d numeric cells are outside the range the template defines' % (
        affected, populated)
    return res


@check('V6', 'Clean text (no stray spaces)')
def v6(ctx: Context) -> CheckResult:
    res = CheckResult('V6', 'Clean text (no stray spaces)', 'validity')
    coded = {}
    for t in ctx.cmap.mapped_tokens:
        f = ctx.fields.get(t)
        if f is None:
            continue
        ft = field_type(f)
        if ft in ('text', 'cic', 'currency', 'country'):
            coded[t] = ft
    if not coded or not ctx.rows:
        res.detail = 'no text columns mapped'
        res.evaluated = bool(coded)
        return res
    populated = affected = 0
    for i, row in enumerate(ctx.rows):
        for t, ft in coded.items():
            v = ctx.value(row, t)
            if is_blank(v):
                continue
            populated += 1
            s = text_of(v)
            if has_text_noise(v):
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), repr(v)[:60], 'no stray whitespace')
            elif ft in ('cic', 'currency', 'country') and s != s.upper():
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, 'upper-case code')
            elif isinstance(v, float) and v == int(v) and t in ('2', '15'):
                affected += 1
                res.add(ctx.sheet_row(i), ctx.column_name(t), v, 'a code, not a decimal number')
    res.n_applicable, res.n_affected = populated, affected
    res.f = share(affected, populated)
    res.detail = '%d of %d text cells carry stray whitespace or the wrong case' % (affected, populated)
    return res
