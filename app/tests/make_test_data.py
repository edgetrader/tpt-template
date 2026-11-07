#!/usr/bin/env python3
"""Build the test fixtures for tpt_score.

Every file is tiny (3 positions) and carries exactly one deliberate fault, so a
failing test points at one check. The clean file is built to pass all 24
checks: mandatory fields are filled, conditional fields are filled whenever the
template says the condition applies, identifiers carry valid check digits,
weights add up to 100% and the market values add up to the NAV.

    python tests/make_test_data.py [--template <V7 workbook>] [--out tests/data]

Writes the fixtures plus ``expected.json``, which test_tpt_score.py reads.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tpt_score.rules import evaluate_condition, field_type, load_fieldset  # noqa: E402

VALUATION_DATE = '2026-09-30'
REPORTING_DATE = '2026-09-30'
NAV = 10_000_000.0
FUND_ID = 'TESTFUND0001'
FUND_NAME = 'Test Private Markets Fund'
FUND_CCY = 'EUR'

# Three positions: value in portfolio currency, quotation currency, CIC, name.
POSITIONS = [
    {'mv_pc': 5_000_000.0, 'ccy': 'EUR', 'fx': 1.0, 'cic': 'XL89', 'name': 'Test Direct Equity A',
     'country': 'LU', 'isin_seed': 'LU000TEST01'},
    {'mv_pc': 2_000_000.0, 'ccy': 'USD', 'fx': 0.9, 'cic': 'US22', 'name': 'Test Corporate Bond B',
     'country': 'US', 'isin_seed': 'US000TEST02'},
    {'mv_pc': 2_000_000.0, 'ccy': 'EUR', 'fx': 1.0, 'cic': 'XL73', 'name': 'Test Cash Deposit C',
     'country': 'DE', 'isin_seed': 'DE000TEST03'},
    {'mv_pc': 1_000_000.0, 'ccy': 'EUR', 'fx': 1.0, 'cic': 'XL41', 'name': 'Test Fund Holding D',
     'country': 'IE', 'isin_seed': 'IE000TEST04'},
]


# --------------------------------------------------------------- identifiers
def isin_with_check_digit(body11: str) -> str:
    """Append the Luhn check digit so the code validates."""
    s = body11.upper()
    digits = ''.join(str(ord(c) - 55) if c.isalpha() else c for c in s)
    total, double = 0, True
    for ch in reversed(digits):
        d = int(ch)
        if double:
            d *= 2
            if d > 9:
                d -= 9
        total += d
        double = not double
    return s + str((10 - total % 10) % 10)


def lei_with_check_digits(body18: str) -> str:
    """Append the ISO 17442 mod-97-10 check digits."""
    s = body18.upper()
    digits = ''.join(str(ord(c) - 55) if c.isalpha() else c for c in s)
    remainder = int(digits + '00') % 97
    return s + '%02d' % (98 - remainder)


# --------------------------------------------------------------- clean table
def _base_value(token: str, field, pos: dict, index: int):
    """A plausible, valid value for one field on one position."""
    fixed = {
        '1': FUND_ID, '2': '99', '3': FUND_NAME, '4': FUND_CCY, '5': NAV,
        '6': VALUATION_DATE, '7': REPORTING_DATE, '8': 100.0, '8b': 100000.0,
        '9': 0.0, '10': 0.0, '11': 'N',
        '1000': 'V7.0 (official) dated 25 November 2024',
    }
    if token in fixed:
        return fixed[token]

    mv_pc = pos['mv_pc']
    mv_qc = round(mv_pc / pos['fx'], 2)
    per_position = {
        '12': pos['cic'],
        '14': isin_with_check_digit(pos['isin_seed']),
        '15': '1',
        '17': pos['name'],
        '17b': 'A',
        '18': 1000.0,
        '19': mv_qc,
        '21': pos['ccy'],
        '22': mv_qc,
        '23': 0.0,
        '24': mv_pc,
        '25': 0.0,
        '26': round(mv_pc / NAV, 6),
        '27': 0.0,
        '28': round(mv_pc / NAV, 6),
        '30': round(mv_pc / NAV, 6),
        '46': 'Test Issuer %d' % (index + 1),
        '47': lei_with_check_digits('TESTLEI00000000%03d' % (index + 1)),
        '52': pos['country'],
        '115': 'N', '116': 'N', '117': 'N', '118': 'N',
        '119': lei_with_check_digits('TESTFUNDISSUER0001'),
        '120': '0', '121': 'N', '122': pos['country'],
        '123': 'Test Fund Issuer', '123a': 'LU',
        '131': '1',
    }
    if token in per_position:
        return per_position[token]

    kind = field_type(field)
    if kind in ('number', 'integer'):
        return 0.0
    if kind == 'date':
        return VALUATION_DATE
    if kind == 'currency':
        return pos['ccy']
    if kind == 'country':
        return pos['country']
    if kind == 'cic':
        return pos['cic']
    return 'N'


def build_clean(fields):
    """Header and rows of the clean fixture, with conditionals filled in."""
    header = [f.name for f in fields.fields]
    index_of = {f.token: i for i, f in enumerate(fields.fields)}

    rows = []
    for n, pos in enumerate(POSITIONS):
        row = [None] * len(header)
        for f in fields.fields:
            if f.flag == 'M':
                row[index_of[f.token]] = _base_value(f.token, f, pos, n)
        rows.append(row)

    # Second pass: fill every conditional field whose condition the rules can
    # read and that comes out true for the row as it now stands.
    for n, (row, pos) in enumerate(zip(rows, POSITIONS)):
        for f in fields.fields:
            if f.flag != 'C':
                continue
            verdict = f.applies_to_cic(pos['cic'][2])
            if isinstance(verdict, str):
                verdict = evaluate_condition(verdict, row, index_of, pos['cic'])
            if verdict is True:
                row[index_of[f.token]] = _base_value(f.token, f, pos, n)
    return header, rows


# --------------------------------------------------------------- fixtures
def write(path, header, rows, sheet='Data', pre=None, post=None, left=0, state='visible'):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.sheet_state = state
    grid = list(pre or []) + [list(header)] + [list(r) for r in rows] + list(post or [])
    for i, row in enumerate(grid, start=1):
        for j, v in enumerate(row, start=1 + left):
            if v is not None:
                ws.cell(i, j, v)
    wb.save(path)
    return path


def col(header, token):
    for i, h in enumerate(header):
        if str(h).split('_')[0] == token:
            return i
    raise KeyError(token)


def make_all(template: str, outdir: str):
    fields = load_fieldset(template)
    header, rows = build_clean(fields)
    os.makedirs(outdir, exist_ok=True)
    expected = []

    def fixture(name, note, hdr, rws, rating=None, min_score=None, max_score=None,
                exact_score=None, fails=(), passes=(), gate=None, **kw):
        path = os.path.join(outdir, name)
        write(path, hdr, rws, **kw)
        expected.append({'file': name, 'note': note, 'rating': rating,
                         'min_score': min_score, 'max_score': max_score,
                         'exact_score': exact_score, 'fails': list(fails),
                         'passes': list(passes), 'gate': gate})

    def clone():
        return [list(r) for r in rows]

    # -- the clean baseline ------------------------------------------------
    fixture('clean.xlsx', 'every check passes', header, rows,
            rating='Accept', min_score=99.0,
            passes=['S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'C1', 'C3', 'C4', 'C5',
                    'V1', 'V2', 'V3', 'V4', 'V5', 'V6', 'K1', 'K2', 'K3', 'K4', 'K5', 'K6'])

    # -- structure ---------------------------------------------------------
    fixture('s1_title_rows.xlsx', 'two title rows above the header', header, rows,
            pre=[['Quarterly TPT delivery'], ['Prepared by Test Fund Admin']],
            fails=['S1'], passes=['S3', 'C1', 'V1', 'K2'], rating='Accept')
    fixture('s1_offset_c4.xlsx', 'table starts at C4', header, rows,
            pre=[[], [], []], left=2, fails=['S1'], passes=['S3', 'C1'], rating='Accept')

    h = list(header)
    r = clone()
    i = col(header, '6')
    h2 = h[:i] + h[i + 1:]
    r2 = [row[:i] + row[i + 1:] for row in r]
    fixture('s3_missing_mandatory.xlsx', 'the valuation date column is missing', h2, r2,
            fails=['S3'], gate='G6')

    keep = [i for i, hh in enumerate(header)
            if str(hh).split('_')[0] not in
            ('1', '2', '3', '4', '5', '6', '7', '8', '8b', '11', '15', '17',
             '21', '23', '25', '26', '27', '28', '30', '52')]
    fixture('g5_many_missing.xlsx', 'over half the mandatory columns are missing',
            [header[i] for i in keep], [[row[i] for i in keep] for row in rows],
            gate='G5', max_score=49.0, rating='Reject')

    h = ['NAV' if str(x).startswith('5_') else
         'Fund ID' if str(x).startswith('1_') else
         'Valuation date' if str(x).startswith('6_') else x for x in header]
    fixture('s4_renamed.xlsx', 'three columns carry a GP short label', h, rows,
            fails=['S4'], passes=['S3'], rating='Accept')

    i = col(header, '22')
    fixture('s5_duplicate_column.xlsx', 'the market value column appears twice',
            header + [header[i]], [row + [row[i]] for row in rows],
            fails=['S5'], rating='Accept')

    fixture('s6_extra_columns.xlsx', 'two GP-internal columns added',
            header + ['GP internal ref', 'Comment'],
            [row + ['INT-%d' % n, 'checked'] for n, row in enumerate(rows)],
            fails=['S6'], rating='Accept')

    total = [None] * len(header)
    total[0] = 'TOTAL'
    total[col(header, '24')] = NAV
    fixture('s7_totals_row.xlsx', 'a TOTAL row under the last position', header, rows,
            post=[total], fails=['S7'], passes=['K2', 'K3'], rating='Accept')

    # -- completeness ------------------------------------------------------
    r = clone()
    r[0][col(header, '17')] = None
    r[1][col(header, '46')] = None
    fixture('c1_empty_mandatory.xlsx', 'two mandatory cells left empty', header, r,
            fails=['C1'], rating='Accept')

    # A conditional field that the rules say applies to this row, left empty.
    r = clone()
    cond_token = None
    index_of = {f.token: i for i, f in enumerate(fields.fields)}
    for f in fields.fields:
        if f.flag != 'C':
            continue
        i = index_of[f.token]
        if all(row[i] is not None for row in r):
            cond_token = f.token
            for row in r:
                row[i] = None
            break
    fixture('c2_conditional_empty.xlsx',
            'a conditional field (%s) left empty although its condition applies' % cond_token,
            header, r, fails=['C2'], passes=['C1'], rating='Accept')

    r = clone()
    r[1][col(header, '14')] = None
    fixture('c3_missing_identifier.xlsx', 'one position has no instrument identifier',
            header, r, fails=['C3'], gate='G6')

    r = clone()
    r[1][col(header, '3')] = None
    r[1][col(header, '5')] = None
    fixture('c4_fund_level_gap.xlsx', 'one row has no fund name or NAV', header, r,
            fails=['C4', 'K1'], rating='Accept')

    r = clone()
    for row in r:
        row[col(header, '46')] = 'N/A'
        row[col(header, '17')] = 'TBD'
    fixture('c5_placeholders.xlsx', 'N/A and TBD used instead of values', header, r,
            fails=['C5'], rating='Accept')

    # -- validity ----------------------------------------------------------
    r = clone()
    for row in r:
        row[col(header, '22')] = '%s' % format(row[col(header, '22')], ',.2f')
        row[col(header, '24')] = '%s' % format(row[col(header, '24')], ',.2f')
    fixture('v1_numbers_as_text.xlsx', 'market values stored as text with separators',
            header, r, fails=['V1'], rating='Accept')

    r = clone()
    for row in r:
        row[col(header, '6')] = '30/09/2026'
    fixture('v2_dates_mixed.xlsx', 'valuation date in DD/MM/YYYY, reporting date in ISO',
            header, r, fails=['V2'], rating='Accept')

    r = clone()
    for row in r:
        row[col(header, '21')] = 'Euro'
        row[col(header, '52')] = 'Germany'
    fixture('v3_closed_list_words.xlsx', 'currency and country written as words',
            header, r, fails=['V3'], rating='Accept')

    r = clone()
    r[0][col(header, '14')] = 'LU0000TEST019'[:12]
    r[1][col(header, '14')] = 'US0000TEST02X'[:12]
    fixture('v4_bad_isin.xlsx', 'two ISINs fail their check digit', header, r,
            fails=['V4'], rating='Accept')

    r = clone()
    for row in r:
        row[col(header, '26')] = round(row[col(header, '26')] * 100, 4)
    fixture('v5_percent_scale.xlsx', 'weights given on a 0-100 scale', header, r,
            fails=['V5'], rating='Accept')

    r = clone()
    for row in r:
        row[col(header, '17')] = '  ' + str(row[col(header, '17')]) + ' '
        row[col(header, '46')] = str(row[col(header, '46')]) + '  '
    fixture('v6_stray_spaces.xlsx', 'instrument and issuer names padded with spaces',
            header, r, fails=['V6'], rating='Accept')

    # -- consistency -------------------------------------------------------
    r = clone()
    r[1][col(header, '3')] = 'Test Private Markets Fund (old name)'
    r[1][col(header, '4')] = 'USD'
    fixture('k1_fund_varies.xlsx', 'fund name and currency differ on one row', header, r,
            fails=['K1'], rating='Accept')

    r = clone()
    for row in r:
        row[col(header, '26')] = round(row[col(header, '26')] * 0.9, 6)
    fixture('k2_weights_off.xlsx', 'weights add up to 90%', header, r,
            fails=['K2'], rating='Minor fixes')

    r = clone()
    for row in r:
        row[col(header, '5')] = NAV * 1.3
    fixture('k3_nav_mismatch.xlsx', 'market values are 30% below the stated NAV',
            header, r, fails=['K3'], rating='Accept')

    r = clone()
    r[2][col(header, '22')] = r[2][col(header, '22')] * 2.0
    fixture('k4_fx_inconsistent.xlsx', 'one row implies a different USD rate', header, r,
            fails=['K4'], rating='Accept')

    r = clone()
    r.append(list(r[0]))
    fixture('k5_duplicate_row.xlsx', 'the first position appears twice', header, r,
            fails=['K5'], gate='G6')

    r = clone()
    for row in r:
        row[col(header, '6')] = '2025-12-31'
    fixture('k6_stale_date.xlsx', 'valuation date is nine months before the reporting date',
            header, r, fails=['K6'], rating='Accept')

    # -- gates --------------------------------------------------------------
    v6_keep = [i for i, hh in enumerate(header)
               if not str(hh).split('_')[0] in
               ('140', '141', '142', '143', '144', '145', '146', '147', '148')]
    r = [[row[i] for i in v6_keep] for row in rows]
    h = [header[i] for i in v6_keep]
    i1000 = col(h, '1000')
    for row in r:
        row[i1000] = 'V6.0 (official) dated 2 March 2021'
    fixture('g4_v6_template.xlsx', 'a V6 file delivered as V7', h, r,
            fails=['S2'], gate='G4', max_score=40.0, rating='Reject')

    fixture('g3_header_only.xlsx', 'header but no positions', header, [],
            gate='G3', exact_score=0.0, rating='Reject')

    fixture('g2_not_a_tpt.xlsx', 'an unrelated spreadsheet',
            ['Date', 'Account', 'Amount', 'Comment'],
            [['2026-09-30', 'A-1', 100, 'x'], ['2026-09-30', 'A-2', 200, 'y']],
            gate='G2', exact_score=0.0, rating='Reject')

    # -- regression fixtures ------------------------------------------------
    # Each of these reproduced a defect found during the code review.

    # Excel reports a used range far past the data: a single stray cell 900 rows
    # down once made S7 charge the file for 895 "non-position rows".
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Data'
    for j, h in enumerate(header, start=1):
        ws.cell(1, j, h)
    for i, row in enumerate(rows, start=2):
        for j, v in enumerate(row, start=1):
            if v is not None:
                ws.cell(i, j, v)
    ws.cell(900, 5, ' ')
    wb.save(os.path.join(outdir, 'r1_phantom_rows.xlsx'))
    expected.append({'file': 'r1_phantom_rows.xlsx',
                     'note': 'a stray cell 900 rows below the data must not count as rows',
                     'rating': 'Accept', 'exact_score': 100.0, 'gate': None,
                     'fails': [], 'passes': ['S7'], 'min_score': None, 'max_score': None})

    # A fund reporting a NAV of zero while holding positions once scored 100.
    r = clone()
    for row in r:
        row[col(header, '5')] = 0
    fixture('r2_nav_zero.xlsx', 'NAV reported as zero while positions are held',
            header, r, fails=['K3'])

    # A cover page that happens to be named "Data" once hid the real positions
    # and sent the file to gate G2 with a score of 0.
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Data'
    ws['A1'] = 'Cover page'
    ws['A2'] = 'the positions are on the Positions tab'
    ws2 = wb.create_sheet('Positions')
    for j, h in enumerate(header, start=1):
        ws2.cell(1, j, h)
    for i, row in enumerate(rows, start=2):
        for j, v in enumerate(row, start=1):
            if v is not None:
                ws2.cell(i, j, v)
    wb.save(os.path.join(outdir, 'r3_data_sheet_is_cover.xlsx'))
    expected.append({'file': 'r3_data_sheet_is_cover.xlsx',
                     'note': 'a cover page named Data must not hide the real positions',
                     'rating': 'Accept', 'exact_score': 100.0, 'gate': None,
                     'fails': [], 'passes': ['S3', 'C1'], 'min_score': None, 'max_score': None})

    # A file that is not a workbook at all.
    bad = os.path.join(outdir, 'g1_corrupt.xlsx')
    with open(bad, 'wb') as fh:
        fh.write(b'This is not a workbook, it only has the extension of one.\n')
    expected.append({'file': 'g1_corrupt.xlsx', 'note': 'not a workbook at all',
                     'rating': 'Reject', 'exact_score': 0.0, 'gate': 'G1',
                     'fails': [], 'passes': [], 'min_score': None, 'max_score': None})

    with open(os.path.join(outdir, 'expected.json'), 'w', encoding='utf-8') as fh:
        json.dump(expected, fh, indent=1)
    return expected


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser(description='build the tpt_score test fixtures')
    root = os.path.dirname(here)
    guesses = [os.path.join(root, '_tpl.xlsx'),
               os.path.join(os.path.dirname(root), 'templates', 'TPT_V7  20241125_updated.xlsx'),
               os.path.join(root, 'templates', 'TPT_V7  20241125_updated.xlsx')]
    default = next((g for g in guesses if os.path.isfile(g)), guesses[0])
    ap.add_argument('--template', default=default, help='the FinDatEx V7 workbook')
    ap.add_argument('--out', default=os.path.join(here, 'data'))
    a = ap.parse_args()
    if not os.path.isfile(a.template):
        raise SystemExit('template not found: %s (pass --template)' % a.template)
    exp = make_all(a.template, a.out)
    print('%d fixtures written to %s' % (len(exp), a.out))


if __name__ == '__main__':
    main()
