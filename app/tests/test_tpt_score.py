#!/usr/bin/env python3
"""Tests for tpt_score.

    python tests/test_tpt_score.py              # run everything
    python tests/test_tpt_score.py -v           # with each test named
    python -m unittest tests.test_tpt_score     # same, through unittest

The fixtures in tests/data are built by tests/make_test_data.py; this script
builds them on the fly if they are not there. The last test case is a
calibration run over a folder of synthetic samples and is skipped when that
folder is absent.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

from tpt_score import cli, util                                    # noqa: E402
from tpt_score.checks import CHECKS, CheckResult                   # noqa: E402
from tpt_score.config import Config, load_config, validate         # noqa: E402
from tpt_score.rules import field_type, load_fieldset, token_of    # noqa: E402
from tpt_score.scorer import combine, score_file                   # noqa: E402

DATA = os.path.join(HERE, 'data')
def _find(*names):
    """Look for a file or folder next to the project, then one level up."""
    roots = [ROOT, os.path.dirname(ROOT), os.getcwd()]
    for root in roots:
        for name in names:
            p = os.path.join(root, name)
            if os.path.exists(p):
                return p
    return os.path.join(ROOT, names[0])


def _find_samples():
    """The folder of TPT sample files, wherever it has been put. A folder only
    counts when it actually holds files whose names start with TPT_."""
    for root in (os.path.dirname(ROOT), ROOT, os.getcwd()):
        for name in ('data', 'synthetic_samples_v3', 'samples'):
            p = os.path.join(root, name)
            if os.path.isdir(p) and any(f.startswith('TPT_') and f.endswith('.xlsx')
                                        for f in os.listdir(p)):
                return p
    return ''


TEMPLATE = os.environ.get('TPT_TEMPLATE') or _find(
    '_tpl.xlsx',
    os.path.join('templates', 'TPT_V7  20241125_updated.xlsx'),
    'TPT_V7  20241125_updated.xlsx')
SAMPLES = os.environ.get('TPT_SAMPLES') or _find_samples()

_FIELDS = None
_CFG = Config()


def fields():
    global _FIELDS
    if _FIELDS is None:
        _FIELDS = load_fieldset(TEMPLATE)
    return _FIELDS


def score(name, cfg=None):
    return score_file(os.path.join(DATA, name), fields(), cfg or _CFG)


def setUpModule():
    if not os.path.isfile(TEMPLATE):
        raise unittest.SkipTest('V7 template not found at %s; set TPT_TEMPLATE' % TEMPLATE)
    if not os.path.isfile(os.path.join(DATA, 'expected.json')):
        sys.path.insert(0, HERE)
        import make_test_data
        make_test_data.make_all(TEMPLATE, DATA)


# --------------------------------------------------------------------------- units
class TestValueParsing(unittest.TestCase):
    def test_numbers(self):
        cases = [
            (1234.5, 1234.5, False), ('1234.5', 1234.5, True),
            ('1,234.50', 1234.5, True), ('1.234,50', 1234.5, True),
            ('1 234,50', 1234.5, True), ("1'234.50", 1234.5, True),
            ('EUR 1,234.50', 1234.5, True), ('$1,234.50', 1234.5, True),
            ('(1,234.50)', -1234.5, True), ('48.23%', 48.23, True),
            ('1.5M', 1_500_000.0, True), ('2K', 2000.0, True),
        ]
        for raw, want, from_text in cases:
            got, text = util.as_number(raw)
            self.assertAlmostEqual(got, want, places=6, msg='parsing %r' % raw)
            self.assertEqual(text, from_text, 'text flag for %r' % raw)

    def test_non_numbers(self):
        for raw in (None, '', 'abc', '#DIV/0!', 'N/A', '-', 'TBD'):
            got, _ = util.as_number(raw)
            self.assertIsNone(got, 'should not parse %r' % raw)

    def test_dates(self):
        cases = [
            ('2026-09-30', dt.date(2026, 9, 30), 'ISO'),
            ('2026/09/30', dt.date(2026, 9, 30), 'ISO-slash'),
            ('20260930', dt.date(2026, 9, 30), 'compact'),
            ('30/09/2026', dt.date(2026, 9, 30), 'DMY-slash'),
            ('30.09.2026', dt.date(2026, 9, 30), 'DMY-dot'),
            (dt.datetime(2026, 9, 30), dt.date(2026, 9, 30), 'excel'),
            (46296, dt.date(2026, 10, 1), 'serial'),
        ]
        for raw, want, kind in cases:
            got, got_kind = util.as_date(raw)
            self.assertEqual(got, want, 'parsing %r' % (raw,))
            self.assertEqual(got_kind, kind, 'format name for %r' % (raw,))

    def test_bad_dates(self):
        for raw in ('not a date', '2026-13-45', '', None, 'N/A'):
            got, kind = util.as_date(raw)
            self.assertIsNone(got, 'should not parse %r' % (raw,))
            self.assertEqual(kind, 'none')

    def test_isin_check_digit(self):
        for good in ('US0378331005', 'GB0002634946', 'DE0005190003',
                     'FR0000120271', 'LU0048578792'):
            self.assertTrue(util.valid_isin(good), '%s should be valid' % good)
        for bad in ('US0378331006', 'GB0002634947', 'DE000519000A',
                    'US037833100', '', 'ABC'):
            self.assertFalse(util.valid_isin(bad), '%s should be invalid' % bad)

    def test_lei_check_digits(self):
        for good in ('549300DTUYXVMJXZNY75', '5493001KJTIIGC8Y1R12'):
            self.assertTrue(util.valid_lei(good), '%s should be valid' % good)
        for bad in ('549300DTUYXVMJXZNY76', 'TOOSHORT', ''):
            self.assertFalse(util.valid_lei(bad), '%s should be invalid' % bad)

    def test_cic(self):
        self.assertTrue(util.valid_cic('XL89'))
        self.assertTrue(util.valid_cic('US22'))
        self.assertFalse(util.valid_cic('ZZ99'))        # not an ISO country
        self.assertFalse(util.valid_cic('XL8'))         # too short
        self.assertEqual(util.cic_category('XL89'), '8')

    def test_placeholders_and_errors(self):
        for raw in ('N/A', 'n.a.', '-', 'NULL', 'TBD', 'nil'):
            self.assertTrue(util.is_placeholder(raw))
        self.assertFalse(util.is_placeholder('EUR'))
        self.assertTrue(util.is_excel_error('#REF!'))
        self.assertFalse(util.is_excel_error('REF'))


class TestTemplateRules(unittest.TestCase):
    def test_fieldset_loads(self):
        fs = fields()
        self.assertGreaterEqual(len(fs), 140)
        self.assertIn('1000', fs)
        self.assertIn('8b', fs, 'lettered fields must be part of the field set')
        for t in ('17b', '58b', '94b', '105a', '105b', '123a'):
            self.assertIn(t, fs, 'lettered field %s missing' % t)

    def test_mandatory_fields(self):
        fs = fields()
        for t in ('1', '3', '5', '6', '7', '12', '14', '22', '24', '26'):
            self.assertTrue(fs.get(t).mandatory, 'field %s should be mandatory' % t)

    def test_token_parsing(self):
        self.assertEqual(token_of('22_Market_valuation_in_quotation_currency_(A)'), '22')
        self.assertEqual(token_of('8b_Total_number_of_shares'), '8b')
        self.assertEqual(token_of('105a_SCR_mrkt_FX_up_weight_over_NAV'), '105a')
        self.assertIsNone(token_of('Comment'))

    def test_field_types(self):
        fs = fields()
        self.assertEqual(field_type(fs.get('6')), 'date')
        self.assertEqual(field_type(fs.get('12')), 'cic')
        self.assertEqual(field_type(fs.get('21')), 'currency')
        self.assertEqual(field_type(fs.get('52')), 'country')
        self.assertEqual(field_type(fs.get('22')), 'number')


class TestConfig(unittest.TestCase):
    def test_defaults_are_sound(self):
        self.assertEqual(validate(Config()), [])

    def test_bad_weights_are_reported(self):
        cfg = Config()
        cfg.data['dimensions']['structure'] = 20
        self.assertTrue(any('dimension weights' in p for p in validate(cfg)))

    def test_every_check_has_a_weight(self):
        cfg = Config()
        for cid in CHECKS:
            self.assertGreater(cfg.weight(cid), 0, 'check %s has no weight' % cid)

    def test_rating_bands(self):
        cfg = Config()
        self.assertEqual(cfg.rating(100), 'Accept')
        self.assertEqual(cfg.rating(95), 'Accept')
        self.assertEqual(cfg.rating(94.9), 'Minor fixes')
        self.assertEqual(cfg.rating(80), 'Minor fixes')
        self.assertEqual(cfg.rating(79.9), 'Resubmit')
        self.assertEqual(cfg.rating(50), 'Resubmit')
        self.assertEqual(cfg.rating(49.9), 'Reject')
        self.assertEqual(cfg.rating(0), 'Reject')

    def test_yaml_round_trip(self):
        tmp = tempfile.mkdtemp()
        try:
            path = os.path.join(tmp, 'rules.yaml')
            with open(path, 'w', encoding='utf-8') as fh:
                fh.write('rules_version: test-1\nbands:\n  accept: 90\n')
            cfg = load_config(path)
            self.assertEqual(cfg.rules_version, 'test-1')
            self.assertEqual(cfg.bands['accept'], 90)
            self.assertEqual(cfg.bands['minor_fixes'], 80, 'untouched keys keep their default')
        finally:
            shutil.rmtree(tmp)


# --------------------------------------------------------------------------- scoring
class TestWorkedExample(unittest.TestCase):
    """The example in the proposal, scored from its failure measures alone."""

    EXAMPLE = {
        'S1': 1.000, 'S4': 0.082, 'S7': 0.024,
        'C1': 0.020, 'C2': 0.050, 'C3': 0.030, 'C5': 0.010,
        'V1': 0.080, 'V4': 0.025, 'V6': 0.100,
        'K2': 0.060, 'K5': 0.025,
    }

    def _results(self):
        out = {}
        for cid, spec in CHECKS.items():
            out[cid] = CheckResult(cid, spec.title, spec.dimension,
                                   f=self.EXAMPLE.get(cid, 0.0))
        return out

    def test_dimension_scores(self):
        checks, dims, total = combine(self._results(), Config())
        self.assertAlmostEqual(dims['structure'].score, 83.53, places=2)
        self.assertAlmostEqual(dims['completeness'].score, 94.05, places=2)
        self.assertAlmostEqual(dims['validity'].score, 97.025, places=3)
        self.assertAlmostEqual(dims['consistency'].score, 96.00, places=2)

    def test_total(self):
        _checks, _dims, total = combine(self._results(), Config())
        self.assertAlmostEqual(total, 93.852, places=3)
        self.assertEqual(Config().rating(total), 'Minor fixes')

    def test_critical_floor_applies(self):
        checks, _dims, _total = combine(self._results(), Config())
        self.assertAlmostEqual(checks['C3'].f, 0.25, places=6,
                               msg='C3 is critical, so f is floored at 0.25')
        self.assertAlmostEqual(checks['C3'].raw_f, 0.030, places=6)
        self.assertAlmostEqual(checks['K5'].f, 0.25, places=6)
        self.assertAlmostEqual(checks['S1'].f, 1.0, places=6,
                               msg='a non-critical check keeps its own f')

    def test_points_lost_are_on_the_total_scale(self):
        checks, _dims, total = combine(self._results(), Config())
        self.assertAlmostEqual(checks['S1'].points_lost, 2.25, places=3)
        self.assertAlmostEqual(checks['C3'].points_lost, 1.125, places=3)
        self.assertAlmostEqual(100.0 - sum(c.points_lost for c in checks.values()),
                               total, places=6,
                               msg='the points lost must add up to the score')

    def test_fixing_the_critical_items_reaches_accept(self):
        results = self._results()
        results['C3'].f = 0.0
        results['K5'].f = 0.0
        _checks, _dims, total = combine(results, Config())
        self.assertGreaterEqual(total, 95.0)
        self.assertEqual(Config().rating(total), 'Accept')


class TestCleanFile(unittest.TestCase):
    def test_clean_file_scores_full_marks(self):
        s = score('clean.xlsx')
        failed = sorted(c.id for c in s.checks.values() if c.f > 0)
        self.assertEqual(failed, [], 'no check may fail on the clean file')
        self.assertEqual(s.gates, [])
        self.assertAlmostEqual(s.score, 100.0, places=6)
        self.assertEqual(s.rating, 'Accept')

    def test_clean_file_is_fully_mapped(self):
        s = score('clean.xlsx')
        self.assertEqual(s.version, 'V7')
        self.assertEqual(s.positions, 4)
        self.assertGreaterEqual(s.mapped_columns, 150)
        self.assertEqual(s.header_cell, 'A1')

    def test_no_check_is_left_unevaluated(self):
        s = score('clean.xlsx')
        skipped = sorted(c.id for c in s.checks.values() if c.status == 'Not evaluated')
        self.assertEqual(skipped, [], 'every check should run on a complete file')


class TestFixtureExpectations(unittest.TestCase):
    """Each fixture carries one fault; the matching check must catch it."""

    def setUp(self):
        with open(os.path.join(DATA, 'expected.json'), encoding='utf-8') as fh:
            self.expected = json.load(fh)

    def test_every_fixture_behaves_as_recorded(self):
        for e in self.expected:
            with self.subTest(fixture=e['file'], note=e['note']):
                s = score(e['file'])
                for cid in e['fails']:
                    self.assertGreater(s.checks[cid].f, 0,
                                       '%s should fail in %s' % (cid, e['file']))
                for cid in e['passes']:
                    self.assertEqual(s.checks[cid].f, 0,
                                     '%s should pass in %s' % (cid, e['file']))
                if e['gate']:
                    self.assertTrue(any(g.startswith(e['gate']) for g in s.gates),
                                    'gate %s should fire for %s (got %s)'
                                    % (e['gate'], e['file'], s.gates))
                if e['exact_score'] is not None:
                    self.assertAlmostEqual(s.score, e['exact_score'], places=6)
                if e['min_score'] is not None:
                    self.assertGreaterEqual(s.score, e['min_score'])
                if e['max_score'] is not None:
                    self.assertLessEqual(s.score, e['max_score'])
                if e['rating']:
                    self.assertEqual(s.rating, e['rating'])

    def test_every_check_is_covered_by_a_fixture(self):
        covered = set()
        for e in self.expected:
            covered |= set(e['fails'])
        missing = sorted(set(CHECKS) - covered)
        self.assertEqual(missing, [], 'no fixture exercises these checks: %s' % missing)


class TestGates(unittest.TestCase):
    def test_g1_unreadable_file(self):
        s = score('g1_corrupt.xlsx')
        self.assertEqual(s.score, 0.0)
        self.assertEqual(s.rating, 'Reject')
        self.assertTrue(s.error)

    def test_g2_not_a_tpt_file(self):
        s = score('g2_not_a_tpt.xlsx')
        self.assertEqual(s.score, 0.0)
        self.assertTrue(s.gates[0].startswith('G2'))

    def test_g3_no_positions(self):
        s = score('g3_header_only.xlsx')
        self.assertEqual(s.score, 0.0)
        self.assertTrue(s.gates[0].startswith('G3'))

    def test_g4_wrong_version_caps_at_40(self):
        s = score('g4_v6_template.xlsx')
        self.assertEqual(s.version, 'V6')
        self.assertLessEqual(s.score, 40.0)
        self.assertEqual(s.rating, 'Reject')

    def test_g5_many_missing_mandatory_caps_at_49(self):
        s = score('g5_many_missing.xlsx')
        self.assertLessEqual(s.score, 49.0)
        self.assertTrue(any(g.startswith('G5') for g in s.gates))

    def test_g6_critical_failure_cannot_reach_accept(self):
        for name in ('s3_missing_mandatory.xlsx', 'c3_missing_identifier.xlsx',
                     'k5_duplicate_row.xlsx'):
            with self.subTest(fixture=name):
                s = score(name)
                self.assertLessEqual(s.score, 94.0)
                self.assertNotEqual(s.rating, 'Accept')
                self.assertTrue(any(g.startswith('G6') for g in s.gates))

    def test_caps_are_configurable(self):
        cfg = Config()
        cfg.data['gates']['wrong_version_cap'] = 10
        s = score('g4_v6_template.xlsx', cfg)
        self.assertLessEqual(s.score, 10.0)


class TestStructuralRecovery(unittest.TestCase):
    """A table that is not in A1 is still read, and the content still checked."""

    def test_title_rows_do_not_hide_the_content(self):
        clean, offset = score('clean.xlsx'), score('s1_title_rows.xlsx')
        self.assertEqual(offset.positions, clean.positions)
        self.assertEqual(offset.mapped_columns, clean.mapped_columns)
        self.assertGreater(offset.checks['S1'].f, 0)
        for cid in ('C1', 'V1', 'K2', 'K3'):
            self.assertEqual(offset.checks[cid].f, 0,
                             '%s should still pass after recovery' % cid)

    def test_offset_table_is_found(self):
        s = score('s1_offset_c4.xlsx')
        self.assertEqual(s.header_cell, 'C4')
        self.assertEqual(s.positions, 4)

    def test_recovery_is_logged(self):
        s = score('s1_offset_c4.xlsx')
        self.assertTrue(any('column C' in n for n in s.recovery), s.recovery)

    def test_totals_row_is_not_counted_as_a_position(self):
        s = score('s7_totals_row.xlsx')
        self.assertEqual(s.positions, 4, 'the TOTAL row must not be read as a position')
        self.assertGreater(s.checks['S7'].f, 0)

    def test_renamed_columns_are_still_mapped(self):
        s = score('s4_renamed.xlsx')
        self.assertGreater(s.checks['S4'].f, 0)
        self.assertEqual(s.checks['S3'].f, 0, 'short GP labels should still map')
        self.assertEqual(s.positions, 4)


class TestScoringProperties(unittest.TestCase):
    def test_repeatable(self):
        a, b = score('v1_numbers_as_text.xlsx'), score('v1_numbers_as_text.xlsx')
        self.assertEqual(a.score, b.score)
        self.assertEqual(a.rating, b.rating)

    def test_monotonic(self):
        """Every fixture carries a fault the clean file does not, so none of
        them may score above the clean file."""
        clean = score('clean.xlsx').score
        for name in sorted(os.listdir(DATA)):
            if not name.endswith('.xlsx') or name == 'clean.xlsx':
                continue
            with self.subTest(fixture=name):
                self.assertLessEqual(score(name).score, clean + 1e-9,
                                     '%s scores above the clean file' % name)

    def test_score_is_within_range(self):
        for name in sorted(os.listdir(DATA)):
            if name.endswith('.xlsx'):
                s = score(name)
                self.assertGreaterEqual(s.score, 0.0)
                self.assertLessEqual(s.score, 100.0)

    def test_points_lost_reconcile_with_the_score(self):
        for name in ('v1_numbers_as_text.xlsx', 'k2_weights_off.xlsx', 's7_totals_row.xlsx'):
            with self.subTest(fixture=name):
                s = score(name)
                lost = sum(c.points_lost for c in s.checks.values())
                self.assertAlmostEqual(100.0 - lost, s.score, places=6)

    def test_severity_ordering(self):
        """A worse fault of the same kind must not score higher."""
        self.assertLess(score('g5_many_missing.xlsx').score,
                        score('s3_missing_mandatory.xlsx').score)

    def test_issues_are_reported_with_a_row_and_column(self):
        s = score('c1_empty_mandatory.xlsx')
        issues = s.checks['C1'].issues
        self.assertTrue(issues, 'C1 should list the empty cells')
        for issue in issues:
            self.assertIsNotNone(issue.row)
            self.assertTrue(issue.column)
            self.assertTrue(issue.expected)


class TestRegressions(unittest.TestCase):
    """Defects found in review. Each test names the behaviour that was wrong."""

    def test_trailing_blank_rows_are_not_charged_as_rows(self):
        """A stray cell far below the data once made S7 charge the file for
        hundreds of 'non-position rows'."""
        s = score('r1_phantom_rows.xlsx')
        self.assertEqual(s.positions, 4)
        self.assertEqual(s.checks['S7'].f, 0.0, s.checks['S7'].detail)
        self.assertAlmostEqual(s.score, 100.0, places=6)

    def test_zero_nav_with_positions_is_a_failure(self):
        """A fund reporting a NAV of zero while holding positions once scored
        100 because the check skipped a falsy NAV."""
        s = score('r2_nav_zero.xlsx')
        self.assertEqual(s.checks['K3'].f, 1.0)
        self.assertTrue(s.checks['K3'].evaluated if hasattr(s.checks['K3'], 'evaluated')
                        else True)
        self.assertLess(s.score, 100.0)

    def test_cover_page_named_data_does_not_hide_the_positions(self):
        """Picking the sheet called "Data" by name alone once sent a whole file
        to gate G2 with a score of 0."""
        s = score('r3_data_sheet_is_cover.xlsx')
        self.assertEqual(s.sheet, 'Positions')
        self.assertEqual(s.positions, 4)
        self.assertEqual(s.gates, [])
        self.assertAlmostEqual(s.score, 100.0, places=6)

    def test_infinity_and_nan_are_not_numbers(self):
        """inf and nan were accepted as valid numbers and would poison every
        sum they reached."""
        for bad in (float('inf'), float('-inf'), float('nan'), '1e400'):
            with self.subTest(value=repr(bad)):
                self.assertIsNone(util.as_number(bad)[0])

    def test_sheet_name_is_never_empty(self):
        """A file name made only of characters Excel forbids produced an empty
        sheet name, which Excel refuses."""
        from tpt_score.reporters.excel import _safe_sheet_name
        used = set()
        for raw in ('???', '', '[]:*?/\\', 'normal'):
            with self.subTest(name=raw):
                out = _safe_sheet_name(raw, used)
                self.assertTrue(0 < len(out) <= 31, 'bad sheet name %r' % out)

    def test_duplicate_file_names_get_distinct_sheets(self):
        from tpt_score.reporters.excel import _safe_sheet_name
        used = set()
        names = [_safe_sheet_name('same', used) for _ in range(4)]
        self.assertEqual(len(set(names)), 4, names)

    def test_duplicate_check_stays_linear(self):
        """Finding duplicates re-scanned every row for each duplicate found,
        which took 56 seconds on a file repeating 1,500 identifiers."""
        import time
        from tpt_score.checks.base import Context
        from tpt_score.checks.consistency import k5

        class _Table:
            header_row = 1
            header = ['14_Identification_code_of_the_instrument', 'name', 'value']

        n_pairs = 1500
        rows, cmap_index = [], {'14': 0}
        for i in range(n_pairs):
            row = ['ID%06d' % i, 'name %d' % i, 100.0]
            rows.append(row)
            rows.append(list(row))

        class _Map:
            index = cmap_index

            def has(self, t):
                return t in cmap_index

            def value(self, row, t):
                i = cmap_index.get(t)
                return None if i is None else row[i]

            def get(self, t):
                return cmap_index.get(t)

        table = _Table()
        table.rows = rows
        ctx = Context(path='', table=table, cmap=_Map(), fields=fields(), cfg=_CFG)
        t0 = time.time()
        res = k5(ctx)
        elapsed = time.time() - t0
        self.assertEqual(res.n_affected, 2 * n_pairs)
        self.assertLess(elapsed, 3.0, 'duplicate detection took %.1fs for %d rows'
                        % (elapsed, len(rows)))

    def test_a_large_file_is_scored_in_reasonable_time(self):
        """The workbook was being opened twice for every file."""
        import time
        t0 = time.time()
        score('clean.xlsx')
        self.assertLess(time.time() - t0, 5.0)


# --------------------------------------------------------------------------- run
class TestEndToEndRun(unittest.TestCase):
    def setUp(self):
        self.out = tempfile.mkdtemp(prefix='tpt_test_')

    def tearDown(self):
        shutil.rmtree(self.out, ignore_errors=True)

    def test_cli_produces_every_report(self):
        rc = cli.run(['--in', DATA, '--out', self.out, '--template', TEMPLATE,
                      '--no-timestamp', '--quiet', '--csv'])
        self.assertIn(rc, (0, 1))
        for name in ('summary.xlsx', 'summary.csv', 'TPT_scorecards.xlsx', 'run_summary.html'):
            path = os.path.join(self.out, name)
            self.assertTrue(os.path.isfile(path), '%s was not written' % name)
            self.assertGreater(os.path.getsize(path), 400, '%s looks empty' % name)

    def test_summary_csv_has_a_row_per_file(self):
        import csv
        cli.run(['--in', DATA, '--out', self.out, '--template', TEMPLATE,
                 '--no-timestamp', '--quiet', '--csv'])
        with open(os.path.join(self.out, 'summary.csv'), encoding='utf-8-sig') as fh:
            rows = list(csv.DictReader(fh))
        n_files = len([f for f in os.listdir(DATA) if f.endswith('.xlsx')])
        self.assertEqual(len(rows), n_files)
        self.assertTrue(all(r['rating'] for r in rows))

    def test_summary_workbook_has_remarks_and_gate_codes(self):
        import openpyxl
        cli.run(['--in', DATA, '--out', self.out, '--template', TEMPLATE,
                 '--no-timestamp', '--quiet'])
        self.assertFalse(os.path.exists(os.path.join(self.out, 'summary.csv')))
        wb = openpyxl.load_workbook(os.path.join(self.out, 'summary.xlsx'))
        self.assertEqual(len(wb.sheetnames), 3)
        cells = [str(c.value) for row in wb.active.iter_rows() for c in row if c.value is not None]
        for gate in ('G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7'):
            self.assertIn(gate, cells)
        for head in ('Tab found in', 'Missing columns', 'Extra columns', 'Mandatory columns with blanks',
                     'Wrong data type', 'Date formats (date columns)', 'Remarks'):
            self.assertIn(head, cells)
        self.assertIn('Column findings', wb.sheetnames)

    def test_workbook_has_the_expected_sheets(self):
        import openpyxl
        cli.run(['--in', DATA, '--out', self.out, '--template', TEMPLATE,
                 '--no-timestamp', '--quiet'])
        wb = openpyxl.load_workbook(os.path.join(self.out, 'TPT_scorecards.xlsx'))
        for sheet in ('Run summary', 'Dimension scores', 'Check matrix', 'Issue log', 'Feedback'):
            self.assertIn(sheet, wb.sheetnames)
        self.assertGreater(len(wb.sheetnames), 5, 'per-file scorecards are missing')
        wb.close()

    def test_html_summary_is_self_contained(self):
        cli.run(['--in', DATA, '--out', self.out, '--template', TEMPLATE,
                 '--no-timestamp', '--quiet'])
        with open(os.path.join(self.out, 'run_summary.html'), encoding='utf-8') as fh:
            html = fh.read()
        self.assertIn('<style>', html, 'the stylesheet must be inline')
        self.assertNotIn('<script', html)
        self.assertNotIn('src="http', html, 'nothing may be fetched from outside')

    def test_single_file_input(self):
        rc = cli.run(['--in', os.path.join(DATA, 'clean.xlsx'), '--out', self.out,
                      '--template', TEMPLATE, '--no-timestamp', '--quiet'])
        self.assertEqual(rc, 0, 'a clean file should give a zero exit code')

    def test_missing_folder_is_reported(self):
        rc = cli.run(['--in', os.path.join(self.out, 'nope'), '--out', self.out,
                      '--template', TEMPLATE, '--quiet'])
        self.assertEqual(rc, 2)

    def test_lock_files_are_skipped(self):
        shutil.copy(os.path.join(DATA, 'clean.xlsx'), os.path.join(self.out, 'real.xlsx'))
        shutil.copy(os.path.join(DATA, 'clean.xlsx'), os.path.join(self.out, '~$lock.xlsx'))
        found = cli.find_files(self.out)
        self.assertEqual([os.path.basename(f) for f in found], ['real.xlsx'])


@unittest.skipUnless(SAMPLES and os.path.isdir(SAMPLES),
                     'sample folder not found; set TPT_SAMPLES to run the calibration')
class TestCalibration(unittest.TestCase):
    """The acceptance criteria from the proposal, against the sample set."""

    @classmethod
    def setUpClass(cls):
        cls.scores = {}
        for name in sorted(os.listdir(SAMPLES)):
            if name.endswith('.xlsx') and name.startswith('TPT_'):
                cls.scores[name] = score_file(os.path.join(SAMPLES, name), fields(), _CFG)

    def test_clean_baselines_are_accepted(self):
        clean = {n: s for n, s in self.scores.items() if 'CLEAN' in n}
        self.assertGreaterEqual(len(clean), 5, 'no clean baselines found')
        for name, s in clean.items():
            with self.subTest(file=name):
                self.assertGreaterEqual(s.score, 95.0)
                self.assertEqual(s.rating, 'Accept')

    def test_wrong_version_files_are_rejected(self):
        for name, s in self.scores.items():
            if '_C10v' in name or '_C11v' in name:        # V6 and V8 layouts
                with self.subTest(file=name):
                    self.assertLessEqual(s.score, 40.0)
                    self.assertEqual(s.rating, 'Reject')

    def test_unusable_files_score_zero(self):
        for name, s in self.scores.items():
            if '_L08v' in name or '_F05v1' in name or '_F06v' in name:
                with self.subTest(file=name):
                    self.assertEqual(s.score, 0.0, '%s should be unusable' % name)

    def test_problem_files_score_below_the_clean_ones(self):
        clean = [s.score for n, s in self.scores.items() if 'CLEAN' in n]
        problems = [s.score for n, s in self.scores.items() if 'CLEAN' not in n]
        self.assertLess(sum(problems) / len(problems), sum(clean) / len(clean))

    def test_run_is_quick_enough(self):
        import time
        t0 = time.time()
        for name in list(self.scores)[:20]:
            score_file(os.path.join(SAMPLES, name), fields(), _CFG)
        per_file = (time.time() - t0) / 20
        self.assertLess(per_file, 1.8, 'a 100-file run must finish in about three minutes')


class TestRemarks(unittest.TestCase):
    """Plain-language remarks and the gate reference."""

    cfg = _CFG

    def setUp(self):
        self.out_dir = tempfile.mkdtemp(prefix='tpt_rem_')
        self.addCleanup(shutil.rmtree, self.out_dir, ignore_errors=True)

    def _score(self, name):
        return score(name)

    def test_blank_mandatory_columns_are_counted(self):
        fs = self._score('c1_empty_mandatory.xlsx')
        self.assertTrue(fs.blank_mandatory)
        name, empty, total = fs.blank_mandatory[0]
        self.assertGreater(empty, 0)
        self.assertLessEqual(empty, total)

    def test_tab_name_is_recorded(self):
        self.assertTrue(self._score('clean.xlsx').sheet)

    def test_clean_file_has_no_issues(self):
        self.assertEqual(self._score('clean.xlsx').remarks, 'No issues found.')

    def test_missing_columns_are_named(self):
        fs = self._score('s3_missing_mandatory.xlsx')
        self.assertTrue(fs.missing_columns)
        from tpt_score.remarks import column_findings
        self.assertTrue(any(i.endswith('(mandatory)') for i in column_findings(fs)['missing']))
        self.assertIn(fs.missing_columns[0].replace('_', ' ') + ' (mandatory)', column_findings(fs)['missing'])

    def test_renamed_columns_show_both_names(self):
        fs = self._score('s4_renamed.xlsx')
        self.assertTrue(fs.renamed_columns)
        found, expected = fs.renamed_columns[0]
        from tpt_score.remarks import column_findings
        self.assertTrue(any(expected.replace('_', ' ') in i and '->' in i
                            for i in column_findings(fs)['renamed']))

    def test_offset_header_is_reported(self):
        self.assertIn('not A1', self._score('s1_offset_c4.xlsx').remarks)

    def test_wrong_version_and_unreadable(self):
        self.assertIn('not V7', self._score('g4_v6_template.xlsx').remarks)
        self.assertIn('Could not be opened', self._score('g1_corrupt.xlsx').remarks)

    def test_different_template_is_highlighted_in_remarks(self):
        from tpt_score.remarks import template_problems
        fs = self._score('g4_v6_template.xlsx')
        self.assertTrue(template_problems(fs))
        self.assertTrue(fs.remarks.startswith('DIFFERENT TEMPLATE'))
        self.assertEqual(template_problems(self._score('clean.xlsx')), [])

    def test_second_table_on_another_sheet_rejects(self):
        fs = self._score('m1_two_sheets.xlsx')
        self.assertEqual(len(fs.extra_tables), 1)
        self.assertEqual(fs.rating, 'Reject')
        self.assertTrue(any(g.startswith('G7') for g in fs.gates))
        self.assertIn('MORE THAN ONE TPT TABLE', fs.remarks)

    def test_second_block_on_same_sheet_rejects(self):
        fs = self._score('m2_second_block.xlsx')
        self.assertEqual(len(fs.extra_tables), 1)
        self.assertEqual(fs.rating, 'Reject')

    def test_repeated_header_under_data_is_not_a_second_table(self):
        self.assertEqual(self._score('s7_totals_row.xlsx').extra_tables, [])
        self.assertEqual(self._score('clean.xlsx').extra_tables, [])

    def test_wrong_data_type_is_listed_by_column(self):
        from tpt_score.remarks import column_findings
        fs = self._score('t1_text_in_number_column.xlsx')
        self.assertTrue(fs.type_mismatches)
        name, expected, bad, total, example, why = fs.type_mismatches[0]
        self.assertIn('number', expected)
        self.assertEqual(bad, 1)
        self.assertEqual(example, 'about 1m')
        self.assertTrue(column_findings(fs)['types'])
        self.assertEqual(self._score('clean.xlsx').type_mismatches, [])

    def test_date_formats_are_listed_for_every_date_column(self):
        from tpt_score.remarks import dates_standard
        clean = self._score('clean.xlsx')
        self.assertGreaterEqual(len(clean.date_formats), 3)
        self.assertTrue(dates_standard(clean))
        mixed = self._score('t2_mixed_date_formats.xlsx')
        self.assertFalse(dates_standard(mixed))
        formats = {f for col in mixed.date_formats for f, _ in col[1]}
        self.assertIn('DD/MM/YYYY', formats)
        self.assertIn('YYYY-MM-DD', formats)

    def test_accepted_files_list_has_full_paths(self):
        cli.run(['--in', DATA, '--out', self.out_dir, '--template', TEMPLATE, '--no-timestamp', '--quiet'])
        with open(os.path.join(self.out_dir, 'accepted_files.txt'), encoding='utf-8') as fh:
            lines = [l.strip() for l in fh if l.strip()]
        self.assertTrue(lines)
        self.assertTrue(all(os.path.isabs(l) and os.path.isfile(l) for l in lines))
        self.assertIn(os.path.abspath(os.path.join(DATA, 'clean.xlsx')), lines)
        self.assertNotIn(os.path.abspath(os.path.join(DATA, 'g1_corrupt.xlsx')), lines)
        self.assertNotIn(os.path.abspath(os.path.join(DATA, 'm1_two_sheets.xlsx')), lines)

    def test_gate_reference_covers_all_gates(self):
        from tpt_score.remarks import gate_reference
        rows = gate_reference(self.cfg)
        self.assertEqual([r[0] for r in rows], ['G1', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7'])
        self.assertIn('S3', rows[5][1])

    def test_reports_carry_remarks_and_gates(self):
        import csv, tempfile
        from tpt_score.reporters.console import write_csv
        from tpt_score.reporters.html import write_html
        scores = [self._score('clean.xlsx'), self._score('g4_v6_template.xlsx')]
        with tempfile.TemporaryDirectory() as d:
            write_csv(os.path.join(d, 's.csv'), scores)
            rows = list(csv.DictReader(open(os.path.join(d, 's.csv'), encoding='utf-8-sig')))
            self.assertIn('remarks', rows[0])
            self.assertIn('date_formats', rows[0])
            self.assertIn('missing_columns', rows[0])
            write_html(os.path.join(d, 'r.html'), scores, self.cfg, {})
            page = open(os.path.join(d, 'r.html'), encoding='utf-8').read()
            self.assertIn('Gates: hard rules', page)
            self.assertIn('not V7', page)


class TestManagementSummary(unittest.TestCase):
    """The one-page management summary."""

    def setUp(self):
        self.out = tempfile.mkdtemp(prefix='tpt_mgmt_')
        self.addCleanup(shutil.rmtree, self.out, ignore_errors=True)

    def _run(self, *extra):
        cli.run(['--in', DATA, '--out', self.out, '--template', TEMPLATE, '--quiet'] + list(extra))
        runs = sorted(d for d in os.listdir(self.out) if d.startswith('run_'))
        return os.path.join(self.out, runs[-1])

    def test_page_and_stats_are_written(self):
        d = self._run()
        page = open(os.path.join(d, 'management_summary.html'), encoding='utf-8').read()
        for text in ('Usable files', 'Right first time', 'What is holding files back',
                     'To be resent by the GP', 'no earlier run to compare'):
            self.assertIn(text, page)
        self.assertNotIn('<script', page.lower())
        stats = json.load(open(os.path.join(d, 'run_stats.json'), encoding='utf-8'))
        self.assertEqual(stats['kpis']['files'], len(stats['files']))

    def test_second_run_compares_with_the_first(self):
        import time
        self._run()
        time.sleep(1.1)
        d = self._run()
        page = open(os.path.join(d, 'management_summary.html'), encoding='utf-8').read()
        self.assertIn('unchanged since last run', page)
        self.assertIn('also failed last run', page)

    def test_every_check_and_gate_has_a_theme(self):
        from tpt_score.reporters.management import CHECK_THEME, GATE_THEME, THEME_LABEL
        self.assertEqual(set(CHECK_THEME), set(CHECKS))
        for theme in list(CHECK_THEME.values()) + list(GATE_THEME.values()):
            self.assertIn(theme, THEME_LABEL)

    def test_gp_map_groups_funds(self):
        path = os.path.join(self.out, 'gp.csv')
        with open(path, 'w', encoding='utf-8') as fh:
            fh.write('pattern,gp\n*,Example GP\n')
        d = self._run('--gp-map', path)
        page = open(os.path.join(d, 'management_summary.html'), encoding='utf-8').read()
        self.assertIn('By GP, weakest first', page)
        self.assertIn('Example GP', page)

    def test_main_problem_of_a_wrong_template_file(self):
        from tpt_score.reporters.management import main_problem
        self.assertEqual(main_problem(score('g4_v6_template.xlsx')), 'template')
        self.assertEqual(main_problem(score('m1_two_sheets.xlsx')), 'tables')
        self.assertEqual(main_problem(score('g1_corrupt.xlsx')), 'unusable')


if __name__ == '__main__':
    unittest.main(verbosity=2)
