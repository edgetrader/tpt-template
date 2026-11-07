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
from .reporters import recap, write_csv, write_html, write_workbook
from .rules import load_fieldset
from .scorer import FileScore, score_file

# Reports go to an "output" folder beside the app folder, so the scripts and
# what they produce stay separate: <project>/app/ and <project>/output/.
APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(os.path.dirname(APP_DIR), 'output')


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='tpt_score',
        description='Check and score GP submissions on the FinDatEx TPT V7 template. '
                    'Every file gets a score from 0 (unusable) to 100 (no check failed).',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='example:\n'
               '  python tpt_score.py --in ../submissions/2026-Q3\n'
               '  python tpt_score.py --in ../submissions/2026-Q3 --out ../output\n')
    p.add_argument('-i', '--in', dest='input', required=True, metavar='FOLDER',
                   help='folder holding the TPT files to check (a single file also works)')
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
    p.add_argument('--no-excel', action='store_true', help='skip the Excel workbook')
    p.add_argument('--no-html', action='store_true', help='skip the HTML summary')
    p.add_argument('--no-timestamp', action='store_true',
                   help='write straight into --out instead of a timestamped subfolder')
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
        scores.append(s)
        if not opts.quiet:
            print('  [%*d/%d] %-52s %5.1f  %s'
                  % (len(str(len(files))), n, len(files), s.filename[:52], s.score, s.rating))

    outdir = opts.output
    if not opts.no_timestamp:
        outdir = os.path.join(outdir, started.strftime('run_%Y%m%d_%H%M%S'))
    os.makedirs(outdir, exist_ok=True)

    run_info = {
        'started': started.strftime('%Y-%m-%d %H:%M:%S'),
        'folder': os.path.abspath(opts.input),
        'template': template,
        'config': cfg.source,
        'seconds': round(time.time() - t0, 1),
    }

    outputs = {'csv': write_csv(os.path.join(outdir, 'summary.csv'), scores)}
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
