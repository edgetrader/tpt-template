"""Command line: point the tool at a folder of GP submissions and score them."""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import os
import sys
import time
from typing import List

from . import __version__
from .config import Config, load_config, validate
from .loader import READABLE_EXT
from .reporters import recap, write_csv, write_html, write_accepted_xlsx, write_summary_xlsx, write_workbook
from .rules import load_fieldset
from .scorer import FileScore, score_file

# Reports go to an "output" folder beside the app folder, so the scripts and
# what they produce stay separate: <project>/app/ and <project>/output/.
APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(os.path.dirname(APP_DIR), 'output')

# Every file a run writes; a new run removes them first so nothing stale is left.
REPORT_FILES = ('summary.xlsx', 'TPT_scorecards.xlsx', 'accepted_files.xlsx', 'run_summary.html',
                'management_summary.html', 'summary.csv', 'run_stats.json',
                'tpt_combined.xlsx')


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='tpt_score',
        description='Check and score GP submissions on the FinDatEx TPT V7 template. '
                    'Every file gets a score from 0 (unusable) to 100 (no check failed).',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='example:\n'
               '  python tpt_score.py --in ../submissions/2026-Q3\n'
               '  python tpt_score.py --in ../submissions/2026-Q3 --out ../output\n'
               '  python tpt_score.py            (scores input_folder from config/settings.yaml)\n'
               '  python tpt_score.py --in 2026-Q4   (scores <input_folder>/2026-Q4)\n')
    p.add_argument('-i', '--in', dest='input', default=None, metavar='FOLDER',
                   help='folder holding the TPT files to check (a single file also works). '
                        'A relative path is read from input_folder in config/settings.yaml; '
                        'left out, input_folder itself is scored')
    p.add_argument('-s', '--settings', default=None, metavar='FILE',
                   help='settings file with input_folder and fund_reference '
                        '(default: config/settings.yaml in the project folder, if present)')
    p.add_argument('-o', '--out', dest='output', default=DEFAULT_OUT, metavar='FOLDER',
                   help='where the reports go (default: the output folder beside the app '
                        'folder, currently %s)' % DEFAULT_OUT)
    p.add_argument('-t', '--template', metavar='FILE',
                   help='the FinDatEx V7 workbook the rules are read from')
    p.add_argument('-c', '--config', metavar='FILE',
                   help='rules.yaml with the weights, tolerances and thresholds')
    p.add_argument('-r', '--recursive', action='store_true',
                   help='also score files in subfolders')
    p.add_argument('--pattern', default='*', metavar='GLOB',
                   help='only files matching this pattern (default: every Excel file)')
    p.add_argument('-f', '--fund-ref', default=None, metavar='XLSX',
                   help='fund reference workbook; its "Fund Name" tab lists GP Name and Fund Name '
                        '(default: fund_reference in config/settings.yaml)')
    p.add_argument('--gp-map', default=None, help='CSV of pattern,GP name to group funds by GP')
    p.add_argument('--csv', action='store_true', help='also write summary.csv')
    p.add_argument('--no-excel', action='store_true', help='skip the Excel workbook')
    p.add_argument('--no-html', action='store_true', help='skip the HTML summary')
    p.add_argument('--no-timestamp', action='store_true',
                   help='write straight into --out instead of a subfolder named after the input')
    p.add_argument('-q', '--quiet', action='store_true', help='only print the recap')
    p.add_argument('--version', action='version', version='tpt_score %s' % __version__)
    return p


def find_files(root: str, pattern: str = '*', recursive: bool = False) -> List[str]:
    """Every Excel file under ``root``, sorted. A path to one file returns it."""
    if os.path.isfile(root):
        return [root]
    if not os.path.isdir(root):
        raise FileNotFoundError('input folder not found: %s' % root)
    found = []
    walker = os.walk(root) if recursive else [(root, [], os.listdir(root))]
    for folder, _dirs, names in walker:
        for name in names:
            if name.startswith('~$') or name.startswith('.'):
                continue                      # Excel lock files and hidden files
            if os.path.splitext(name)[1].lower() not in READABLE_EXT:
                continue
            if pattern != '*' and not glob.fnmatch.fnmatch(name, pattern):
                continue
            found.append(os.path.join(folder, name))
    return sorted(found)


def input_label(input_path: str, given: str = None) -> str:
    """The output subfolder for an input: the name of the --in folder itself, so
    --in data gives output/data and --in 2026/Q4 gives output/Q4. A single file is
    filed under the folder it sits in."""
    path = os.path.abspath(input_path)
    if os.path.isfile(path):
        path = os.path.dirname(path)
    name = os.path.basename(os.path.normpath(path))
    return ''.join(ch if ch.isalnum() or ch in ' -_.()' else '_' for ch in name).strip() or 'input'


def resolve_template(cfg: Config, given: str, input_path: str) -> str:
    """Find the V7 workbook: --template, then the config path, tried next to the
    input folder and the working directory."""
    candidates = []
    if given:
        candidates.append(given)
    cfg_template = cfg.data.get('template')
    if cfg_template:
        candidates.append(cfg_template)
        base = input_path if os.path.isdir(input_path) else os.path.dirname(input_path)
        for root in (base, os.path.dirname(os.path.abspath(base)), os.getcwd()):
            candidates.append(os.path.join(root, cfg_template))
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    raise FileNotFoundError(
        'the V7 template workbook was not found. Looked for:\n  ' +
        '\n  '.join(str(c) for c in candidates) +
        '\nPass it with --template.')


def run(args=None) -> int:
    # openpyxl warns about Excel features it does not read (data validation,
    # conditional formats, unknown extensions). They say nothing about the data,
    # so they are silenced; real problems still surface as errors.
    import warnings
    warnings.filterwarnings('ignore')
    opts = build_parser().parse_args(args)
    started = dt.datetime.now()

    try:
        cfg = load_config(opts.config)
    except Exception as exc:
        print('config error: %s' % exc, file=sys.stderr)
        return 2
    problems = validate(cfg)
    if problems:
        print('config error:\n  ' + '\n  '.join(problems), file=sys.stderr)
        return 2

    # Paths not given on the command line come from config/settings.yaml.
    from .settings import SettingsError, load_settings
    try:
        settings = load_settings(opts.settings)
    except SettingsError as exc:
        print('settings error: %s' % exc, file=sys.stderr)
        return 2
    given_in = opts.input
    # input_folder is the root: --in names a folder under it, or the root itself
    # when --in is left out. An absolute --in is used as given.
    if not opts.input:
        opts.input = settings.input_folder
    elif settings.input_folder and not os.path.isabs(os.path.expanduser(opts.input)):
        opts.input = os.path.normpath(os.path.join(settings.input_folder, opts.input))
    if not opts.fund_ref and settings.fund_reference:
        opts.fund_ref = settings.fund_reference
    if not opts.input:
        print('no input folder: pass --in FOLDER or set input_folder in %s'
              % (settings.path or 'config/settings.yaml'), file=sys.stderr)
        return 2
    if settings.path and not opts.quiet:
        print('settings: %s' % settings.path)

    try:
        template = resolve_template(cfg, opts.template, opts.input)
        files = find_files(opts.input, opts.pattern, opts.recursive)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 2

    if not files:
        print('no Excel files found in %s' % opts.input, file=sys.stderr)
        return 1

    if not opts.quiet:
        print('tpt_score %s | rules %s | %d files from %s'
              % (__version__, cfg.rules_version, len(files), opts.input))

    fund_ref = None
    if opts.fund_ref:
        from .fundref import FundReferenceError, load_fund_reference
        try:
            fund_ref = load_fund_reference(opts.fund_ref)
        except FundReferenceError as exc:
            print('error: %s' % exc, file=sys.stderr)
            return 2
        ref_abs = os.path.abspath(opts.fund_ref)
        files = [f for f in files if os.path.abspath(f) != ref_abs]   # never score the reference
        if not opts.quiet:
            print('fund reference: %d funds from %s' % (len(fund_ref), opts.fund_ref))

    fields = load_fieldset(template)

    scores: List[FileScore] = []
    t0 = time.time()
    for n, path in enumerate(files, start=1):
        try:
            s = score_file(path, fields, cfg)
        except Exception as exc:                      # never let one file stop the run
            s = FileScore(path=path, filename=os.path.basename(path), score=0.0,
                          rating='Reject', error='%s: %s' % (type(exc).__name__, str(exc)[:160]),
                          gates=['G1: the file could not be processed'],
                          rules_version=cfg.rules_version)
        if fund_ref is not None:
            # The GP comes from the fund name in the FILE NAME, so it is known
            # even for a file that could not be read.
            m = fund_ref.match_filename(s.filename)
            if True:
                s.fund_ref_status, s.fund_ref_label = m.status, m.label
                s.gp_name, s.ref_fund = m.gp, m.ref_fund
        scores.append(s)
        if not opts.quiet:
            print('  [%*d/%d] %-52s %5.1f  %s'
                  % (len(str(len(files))), n, len(files), s.filename[:52], s.score, s.rating))

    # Reports go to <output>/<input folder name>/. The latest run over a folder
    # replaces the reports of the one before; its figures are read first so the
    # management page can still show the change since then.
    outdir = opts.output
    if not opts.no_timestamp:
        outdir = os.path.join(outdir, input_label(opts.input, given_in))
    os.makedirs(outdir, exist_ok=True)
    from .reporters.management import previous_stats
    prev = previous_stats(outdir, os.path.abspath(opts.input))
    for name in REPORT_FILES:                       # nothing stale from an earlier run
        old = os.path.join(outdir, name)
        if os.path.isfile(old):
            try:
                os.remove(old)
            except OSError:
                pass

    run_info = {
        'started': started.strftime('%Y-%m-%d %H:%M:%S'),
        'folder': os.path.abspath(opts.input),
        'template': template,
        'config': cfg.source,
        'seconds': round(time.time() - t0, 1),
        'fund_reference': os.path.abspath(opts.fund_ref) if opts.fund_ref else '',
    }

    outputs = {'summary': write_summary_xlsx(os.path.join(outdir, 'summary.xlsx'), scores, cfg, run_info)}
    from .reporters.management import (GP_MAP_NAME, load_gp_map,
                                       write_management, write_stats)
    gp_map = load_gp_map([opts.gp_map, os.path.join(os.path.abspath(opts.input) if os.path.isdir(opts.input)
                                                   else os.path.dirname(os.path.abspath(opts.input)), GP_MAP_NAME),
                          os.path.join(APP_DIR, GP_MAP_NAME)])
    outputs['management'] = write_management(os.path.join(outdir, 'management_summary.html'),
                                             scores, cfg, run_info, gp_map, prev)
    write_stats(outdir, scores, run_info)
    outputs['accepted'] = write_accepted_xlsx(os.path.join(outdir, 'accepted_files.xlsx'), scores)
    if opts.csv:
        outputs['csv'] = write_csv(os.path.join(outdir, 'summary.csv'), scores)
    if not opts.no_excel and cfg.data.get('reports', {}).get('excel', True):
        outputs['excel'] = write_workbook(os.path.join(outdir, 'TPT_scorecards.xlsx'),
                                          scores, cfg, run_info)
    if not opts.no_html and cfg.data.get('reports', {}).get('html', True):
        outputs['html'] = write_html(os.path.join(outdir, 'run_summary.html'),
                                     scores, cfg, run_info)

    print()
    print(recap(scores, cfg, outputs))

    # A non-zero exit code when anything needs action, so the run can be chained.
    return 0 if all(s.rating in ('Accept', 'Minor fixes') for s in scores) else 1


if __name__ == '__main__':      # pragma: no cover
    sys.exit(run())
