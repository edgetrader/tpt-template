"""One-page management summary: plain words, grouped by fund or GP, with
change since the previous run.

The analyst reports (summary.xlsx, run_summary.html) say what is wrong with each
file in template terms. This page answers management's questions instead: how
much of what came in can we use, what kinds of problem are holding files back,
which GPs need chasing, and is it getting better.
"""
from __future__ import annotations

import collections
import csv
import fnmatch
import glob
import html
import json
import os
from typing import Dict, List, Optional, Tuple

from ..scorer import FileScore

USABLE = ('Accept', 'Minor fixes')
BLOCKING = ('Resubmit', 'Reject')
BAND_COLOUR = {'Accept': '#2e7d32', 'Minor fixes': '#b58100',
               'Resubmit': '#d4760a', 'Reject': '#c62828'}

# Business themes. Every check and gate belongs to one; the action is what we
# ask the GP to do about it.
THEMES = [
    ('unusable', 'File unusable: unreadable, empty or not a TPT',
     'Resend the file on the V7 template'),
    ('template', 'Wrong template version',
     'Resend on the FinDatEx TPT V7 template'),
    ('tables', 'More than one table in the file',
     'Send one table per file, one fund per table'),
    ('missing', 'Required data missing',
     'Add the missing mandatory columns and values'),
    ('reconcile', 'Figures do not reconcile (weights, NAV, FX)',
     'Check that weights sum to 100% and values add up to the NAV'),
    ('identifiers', 'Invalid security or issuer identifiers',
     'Correct the ISIN and LEI codes'),
    ('format', 'Values in the wrong format',
     'Use numbers, ISO dates and ISO codes as the template requires'),
    ('consistency', 'Inconsistent, duplicated or stale data',
     'Remove duplicates and use one valuation and reporting date'),
    ('optional', 'Conditional or fund-level fields incomplete',
     'Fill the fields the template requires for these positions'),
    ('layout', 'Layout differs from the template (cosmetic)',
     'Start the table at A1 with the template column names'),
]
THEME_LABEL = {k: label for k, label, _ in THEMES}
THEME_ACTION = {k: action for k, _, action in THEMES}

CHECK_THEME = {
    'S1': 'layout', 'S4': 'layout', 'S5': 'layout', 'S6': 'layout', 'S7': 'layout',
    'S2': 'template', 'S3': 'missing', 'C1': 'missing', 'C3': 'missing',
    'C2': 'optional', 'C4': 'optional', 'C5': 'optional',
    'V1': 'format', 'V2': 'format', 'V3': 'format', 'V5': 'format', 'V6': 'format',
    'V4': 'identifiers',
    'K2': 'reconcile', 'K3': 'reconcile', 'K4': 'reconcile',
    'K1': 'consistency', 'K5': 'consistency', 'K6': 'consistency',
}
GATE_THEME = {'G1': 'unusable', 'G2': 'unusable', 'G3': 'unusable', 'G4': 'template',
              'G5': 'missing', 'G7': 'tables'}

# An issue counts as material when it costs at least this many points. Below
# that it is noise for a management reader (it stays in the analyst reports).
MATERIAL_POINTS = 0.5

STATS_FILE = 'run_stats.json'
GP_MAP_NAME = 'gp_map.csv'


# --------------------------------------------------------------------------- grouping

def load_gp_map(paths: List[str]) -> List[Tuple[str, str]]:
    """Read the first gp_map.csv found: rows of (pattern, gp name).

    A pattern is matched, case-insensitive, against the fund ID, the fund name
    and the file name; wildcards (*, ?) are allowed. The first match wins.
    """
    for p in paths:
        if p and os.path.isfile(p):
            rows = []
            with open(p, encoding='utf-8-sig', newline='') as fh:
                for r in csv.reader(fh):
                    if len(r) >= 2 and r[0].strip() and not r[0].strip().startswith('#') \
                            and r[0].strip().lower() not in ('pattern', 'match'):
                        rows.append((r[0].strip().lower(), r[1].strip()))
            return rows
    return []


def _usable_text(v: str) -> bool:
    v = (v or '').strip()
    return bool(v) and v.lower() not in ('text', 'n/a', 'na', 'none', 'null', '-', 'tbc')


def fund_label(s: FileScore) -> str:
    if _usable_text(s.fund_name):
        return s.fund_name.strip()
    if _usable_text(s.fund_id):
        return s.fund_id.strip()
    return '(fund not identified)'


def gp_of(s: FileScore, gp_map: List[Tuple[str, str]]) -> str:
    if s.gp_name:                       # from the fund reference workbook
        return s.gp_name
    keys = [k.lower() for k in (s.fund_id or '', s.fund_name or '', s.filename) if k]
    for pattern, gp in gp_map:
        wild = any(ch in pattern for ch in '*?')
        for k in keys:
            if (wild and fnmatch.fnmatch(k, pattern)) or (not wild and pattern in k):
                return gp
    return ''


# --------------------------------------------------------------------------- themes

def themes_of(s: FileScore) -> List[str]:
    """Material business themes for one file, most serious first."""
    found = []
    for g in s.gates:
        t = GATE_THEME.get(g[:2])
        if t and t not in found:
            found.append(t)
    if any(g.startswith(('G1', 'G2', 'G3')) for g in s.gates):
        return found
    for c in sorted(s.checks.values(), key=lambda c: -c.points_lost):
        if c.points_lost >= MATERIAL_POINTS or (c.critical and c.raw_f > 0):
            t = CHECK_THEME.get(c.id)
            if t and t not in found:
                found.append(t)
    order = [k for k, _, _ in THEMES]
    gate_part = [t for t in found if any(GATE_THEME.get(g[:2]) == t for g in s.gates)]
    rest = sorted((t for t in found if t not in gate_part), key=order.index)
    return gate_part + rest


def main_problem(s: FileScore) -> str:
    t = themes_of(s)
    return t[0] if t else ''


# --------------------------------------------------------------------------- stats

def kpis(scores: List[FileScore]) -> Dict[str, float]:
    n = len(scores)
    usable = sum(1 for s in scores if s.rating in USABLE)
    first = sum(1 for s in scores if s.rating == 'Accept')
    action = sum(1 for s in scores if s.rating in BLOCKING)
    return {
        'files': n,
        'usable_pct': 100.0 * usable / n if n else 0.0,
        'first_time_pct': 100.0 * first / n if n else 0.0,
        'average': sum(s.score for s in scores) / n if n else 0.0,
        'need_action': action,
        'usable': usable,
    }


def write_stats(outdir: str, scores: List[FileScore], run_info: dict) -> str:
    path = os.path.join(outdir, STATS_FILE)
    data = {'started': run_info.get('started', ''), 'folder': run_info.get('folder', ''),
            'kpis': kpis(scores),
            'files': {os.path.abspath(s.path): {'score': round(s.score, 2), 'rating': s.rating,
                                                'fund': fund_label(s)} for s in scores}}
    with open(path, 'w', encoding='utf-8') as fh:
        json.dump(data, fh, indent=1)
    return path


def previous_stats(outdir: str, folder: str) -> Optional[dict]:
    """The latest earlier run over the same input folder, if one was kept."""
    parent = os.path.dirname(os.path.abspath(outdir))
    here = os.path.abspath(outdir)
    best = None
    for p in glob.glob(os.path.join(parent, '*', STATS_FILE)):
        d = os.path.dirname(os.path.abspath(p))
        if d == here:
            continue
        try:
            with open(p, encoding='utf-8') as fh:
                data = json.load(fh)
        except (OSError, ValueError):
            continue
        if os.path.abspath(data.get('folder', '')) != os.path.abspath(folder):
            continue
        if best is None or data.get('started', '') > best.get('started', ''):
            best = data
    return best


# --------------------------------------------------------------------------- page

CSS = """
*{box-sizing:border-box}
body{font:13px/1.45 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#1a1a1a;
     margin:0;padding:24px 28px;background:#fff;max-width:1180px}
h1{font-size:20px;margin:0 0 2px}
.sub{color:#666;font-size:11.5px;margin-bottom:12px}
.headline{font-size:15px;line-height:1.5;margin:0 0 16px;padding:10px 14px;background:#f5f7fa;
          border-left:4px solid #1f3864;border-radius:4px}
.tiles{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:18px}
.tile{border:1px solid #d8d8d8;border-radius:8px;padding:10px 13px}
.tile .n{font-size:26px;font-weight:600;line-height:1.1}
.tile .l{font-size:10.5px;color:#555;text-transform:uppercase;letter-spacing:.04em;margin-top:2px}
.tile .d{font-size:11.5px;margin-top:4px;color:#666}
.up,.tile .d.up{color:#2e7d32}.down,.tile .d.down{color:#c62828}
h2{font-size:13px;margin:0 0 6px;padding-bottom:3px;border-bottom:1px solid #e3e3e3}
section{margin-bottom:16px}
.cols{display:grid;grid-template-columns:1.05fr 1fr;gap:26px;align-items:start}
table{border-collapse:collapse;width:100%;font-size:12px;table-layout:fixed}
th{text-align:left;color:#666;font-weight:600;padding:3px 8px 3px 0;border-bottom:1px solid #e3e3e3}
td{padding:4px 8px 4px 0;border-bottom:1px solid #f1f1f1;vertical-align:top}
.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.track{background:#f1f1f1;border-radius:3px;height:12px;width:100%;margin-top:3px}
.bar{height:12px;border-radius:3px;display:block}
.pill{display:inline-block;padding:1px 7px;border-radius:9px;font-size:10.5px;color:#fff;white-space:nowrap}
.muted{color:#888;font-size:10.5px}
.foot{margin-top:14px;color:#777;font-size:10.5px;border-top:1px solid #e3e3e3;padding-top:7px}
@media print{body{padding:0;font-size:11.5px}.headline{font-size:13.5px}}
"""


def _pill(rating: str) -> str:
    return '<span class="pill" style="background:%s">%s</span>' % (
        BAND_COLOUR.get(rating, '#777'), html.escape(rating))


def _delta(now: float, before: Optional[float], unit: str = '', better_up: bool = True,
           digits: int = 1) -> str:
    if before is None:
        return '<div class="d">no earlier run to compare</div>'
    diff = now - before
    if abs(diff) < 10 ** -digits / 2:
        return '<div class="d">unchanged since last run</div>'
    good = (diff > 0) == better_up
    fmt = '%+.' + str(digits) + 'f%s since last run'
    return '<div class="d %s">%s</div>' % ('up' if good else 'down', fmt % (diff, unit))


def _fmt_pct(x: float) -> str:
    return '%.0f%%' % x


def headline(k: dict, theme_counts: collections.Counter, blocking_themes: collections.Counter,
             prev: Optional[dict]) -> str:
    n = k['files']
    if not n:
        return 'No files were received in this run.'
    parts = ['<b>%d of %d files (%s) can be used</b>' % (k['usable'], n, _fmt_pct(k['usable_pct']))]
    if k['need_action']:
        top = blocking_themes.most_common(1)
        reason = ''
        if top:
            lead = [THEME_LABEL[t].split(':')[0].split(' (')[0].lower()
                    for t, c in blocking_themes.most_common() if c == top[0][1]][:2]
            reason = ', mainly %s' % ' and '.join(lead)
        parts.append('%d must be resent by the GP%s' % (k['need_action'], reason))
    text = '; '.join(parts) + '.'
    if prev:
        before = prev.get('kpis', {}).get('usable_pct')
        if before is not None:
            diff = k['usable_pct'] - before
            if abs(diff) >= 0.5:
                text += ' Usable share is %s %.0f points on the last run.' % (
                    'up' if diff > 0 else 'down', abs(diff))
            else:
                text += ' Usable share is unchanged on the last run.'
    return text


def write_management(path: str, scores: List[FileScore], cfg, run_info: dict,
                     gp_map: List[Tuple[str, str]] = (), prev: Optional[dict] = None) -> str:
    k = kpis(scores)
    pk = (prev or {}).get('kpis', {})

    # themes across the run
    theme_files = collections.Counter()
    theme_blocking = collections.Counter()
    theme_funds: Dict[str, set] = collections.defaultdict(set)
    for s in scores:
        for t in themes_of(s):
            theme_files[t] += 1
            theme_funds[t].add(fund_label(s))
        if s.rating in BLOCKING and main_problem(s):
            theme_blocking[main_problem(s)] += 1

    out = ['<!DOCTYPE html><html><head><meta charset="utf-8">',
           '<title>GP TPT submissions: management summary</title>',
           '<style>%s</style></head><body>' % CSS,
           '<h1>GP TPT submissions: management summary</h1>',
           '<div class="sub">Run of %s &middot; %d files &middot; checked against FinDatEx TPT V7</div>'
           % (html.escape(str(run_info.get('started', ''))), k['files'])]
    line = headline(k, theme_files, theme_blocking, prev)
    if any(s.fund_ref_status for s in scores):
        lost = {(s.fund_name or s.fund_id or '') for s in scores
                if s.fund_ref_status in ('not found', 'ambiguous', 'no name')}
        nfiles = sum(1 for s in scores if s.fund_ref_status in ('not found', 'ambiguous', 'no name'))
        if lost:
            line += (' <span class="down">%d fund name%s (%d file%s) could not be matched to a GP '
                     'in the fund reference.</span>' % (len(lost), '' if len(lost) == 1 else 's',
                                                       nfiles, '' if nfiles == 1 else 's'))
        else:
            line += ' Every fund was matched to its GP.'
    out.append('<p class="headline">%s</p>' % line)

    # ---- KPI tiles ----------------------------------------------------------
    out.append('<div class="tiles">')
    out.append('<div class="tile"><div class="n">%s</div><div class="l">Usable files</div>%s</div>'
               % (_fmt_pct(k['usable_pct']), _delta(k['usable_pct'], pk.get('usable_pct'), ' pts', True, 0)))
    out.append('<div class="tile"><div class="n">%s</div><div class="l">Right first time</div>%s</div>'
               % (_fmt_pct(k['first_time_pct']),
                  _delta(k['first_time_pct'], pk.get('first_time_pct'), ' pts', True, 0)))
    out.append('<div class="tile"><div class="n">%.1f</div><div class="l">Average quality score</div>%s</div>'
               % (k['average'], _delta(k['average'], pk.get('average'), '', True, 1)))
    out.append('<div class="tile"><div class="n">%d</div><div class="l">Files to be resent</div>%s</div>'
               % (k['need_action'], _delta(k['need_action'], pk.get('need_action'), '', False, 0)))
    out.append('</div>')

    # ---- two columns: themes | funds --------------------------------------
    out.append('<div class="cols"><div>')
    out.append('<section><h2>What is holding files back</h2><table>'
               '<colgroup><col><col style="width:62px"><col style="width:66px"></colgroup>'
               '<tr><th>Problem</th><th class="n">Files</th><th class="n">Blocking</th></tr>')
    top = max(theme_files.values()) if theme_files else 1
    shown = [t for t, _, _ in THEMES if theme_files[t]]
    shown.sort(key=lambda t: (-theme_blocking[t], -theme_files[t]))
    for t in shown:
        colour = '#c62828' if theme_blocking[t] else ('#9e9e9e' if t in ('layout', 'optional') else '#b58100')
        out.append('<tr><td>%s<div class="track"><span class="bar" style="width:%.0f%%;background:%s">'
                   '</span></div></td><td class="n">%d</td><td class="n">%s</td></tr>'
                   % (html.escape(THEME_LABEL[t]), 100.0 * theme_files[t] / top, colour,
                      theme_files[t], theme_blocking[t] or '&ndash;'))
    if not shown:
        out.append('<tr><td colspan="3">No material problems in this run.</td></tr>')
    out.append('</table><div class="muted" style="margin-top:4px">Files: files with the problem. '
               'Blocking: files that must be resent with this as their main problem. Only problems that cost at least '
               '%.1f points are counted; minor ones are in the analyst report.</div></section>'
               % MATERIAL_POINTS)
    out.append('</div><div>')

    # funds / GPs
    has_ref = any(s.fund_ref_status for s in scores)
    use_gp = has_ref or (any(gp_of(s, gp_map) for s in scores) if gp_map else False)
    group_name = 'GP' if use_gp else 'Fund'
    groups: Dict[str, List[FileScore]] = collections.defaultdict(list)
    for s in scores:
        key = (gp_of(s, gp_map) or '(GP not identified)') \
            if use_gp else fund_label(s)
        groups[key].append(s)
    rank = {'Reject': 0, 'Resubmit': 1, 'Minor fixes': 2, 'Accept': 3}
    rows = []
    for g, items in groups.items():
        usable = sum(1 for s in items if s.rating in USABLE)
        worst = min(items, key=lambda s: rank.get(s.rating, 0)).rating
        avg = sum(s.score for s in items) / len(items)
        rows.append((g, len(items), usable, avg, worst))
    rows.sort(key=lambda r: (rank.get(r[4], 0), r[2] / r[1], r[3]))
    out.append('<section><h2>By %s, weakest first</h2><table>'
               '<colgroup><col><col style="width:44px"><col style="width:52px">'
               '<col style="width:50px"><col style="width:84px"></colgroup>'
               '<tr><th>%s</th><th class="n">Files</th><th class="n">Usable</th>'
               '<th class="n">Score</th><th>Worst</th></tr>' % ('GP' if use_gp else 'fund', group_name))
    for g, nfiles, usable, avg, worst in rows[:14]:
        out.append('<tr><td>%s</td><td class="n">%d</td><td class="n">%d</td><td class="n">%.0f</td>'
                   '<td>%s</td></tr>' % (html.escape(g), nfiles, usable, avg, _pill(worst)))
    out.append('</table>')
    if len(rows) > 14:
        out.append('<div class="muted" style="margin-top:4px">%d more %ss, all listed in summary.xlsx.</div>'
                   % (len(rows) - 14, 'GP' if use_gp else 'fund'))
    if not use_gp:
        out.append('<div class="muted" style="margin-top:4px">Grouped by fund name. Pass a fund '
                   'reference (--fund-ref) to group by GP.</div>')
    out.append('</section></div></div>')

    # ---- funds the reference does not know ---------------------------------
    if has_ref:
        unknown = collections.OrderedDict()
        for s in sorted(scores, key=lambda x: fund_label(x)):
            if s.fund_ref_status in ('not found', 'ambiguous', 'no name'):
                name = (s.fund_name or s.fund_id or '(no fund name in the file)').strip()
                unknown.setdefault(name, []).append(s)
        close = sorted({(s.fund_name or s.fund_id, s.ref_fund, s.gp_name) for s in scores
                        if s.fund_ref_status == 'close'})
        out.append('<section><h2>Funds not found in the fund reference (%d)</h2>' % len(unknown))
        if unknown:
            out.append('<table><colgroup><col style="width:38%"><col style="width:60px"><col></colgroup>'
                       '<tr><th>Fund name as written by the GP</th><th class="n">Files</th>'
                       '<th>Why</th></tr>')
            for name, items in unknown.items():
                m = items[0]
                why = ('matches more than one fund: ' + m.fund_ref_label.split(': ', 1)[-1]
                       if m.fund_ref_status == 'ambiguous' else
                       'the file holds no fund name' if m.fund_ref_status == 'no name' else
                       'no fund in the reference is close enough')
                out.append('<tr><td><b class="down">%s</b></td><td class="n">%d</td><td>%s</td></tr>'
                           % (html.escape(name), len(items), html.escape(why)))
            out.append('</table>')
        else:
            out.append('<p>Every fund was found.</p>')
        if close:
            out.append('<div class="muted" style="margin-top:4px">%d fund name%s matched only '
                       'approximately; please confirm: %s</div>'
                       % (len(close), '' if len(close) == 1 else 's',
                          html.escape('; '.join('"%s" = %s (%s)' % c for c in close[:6]))
                          + (' and %d more' % (len(close) - 6) if len(close) > 6 else '')))
        out.append('</section>')

    # ---- escalation ---------------------------------------------------------
    esc = sorted((s for s in scores if s.rating in BLOCKING), key=lambda s: (s.score, fund_label(s)))
    prev_files = (prev or {}).get('files', {})
    out.append('<section><h2>To be resent by the GP (%d)</h2>' % len(esc))
    if esc:
        out.append('<table><colgroup><col style="width:24%%"><col style="width:22%%">'
                   '<col style="width:44px"><col style="width:76px"><col style="width:24%%"><col>'
                   '</colgroup><tr><th>%s</th><th>File</th><th class="n">Score</th><th>Rating</th>'
                   '<th>Main problem</th><th>Ask the GP to</th></tr>' % ('GP / fund' if use_gp else 'Fund'))
        for s in esc:
            t = main_problem(s)
            who = fund_label(s)
            if use_gp:
                who = '%s<div class="muted">%s</div>' % (html.escape(gp_of(s, gp_map) or '(GP not identified)'),
                                                         html.escape(who))
            else:
                who = html.escape(who)
            again = prev_files.get(os.path.abspath(s.path), {}).get('rating') in BLOCKING
            out.append('<tr><td>%s%s</td><td style="word-break:break-all">%s</td><td class="n">%.0f</td>'
                       '<td>%s</td><td>%s</td><td>%s</td></tr>'
                       % (who, '<div class="down" style="font-size:10.5px">also failed last run</div>'
                          if again else '', html.escape(s.filename), s.score, _pill(s.rating),
                          html.escape(THEME_LABEL.get(t, 'See analyst report')),
                          html.escape(THEME_ACTION.get(t, 'Review the scorecard'))))
        out.append('</table>')
    else:
        out.append('<p>No file needs to be resent.</p>')
    out.append('</section>')

    out.append('<div class="foot"><b>Ratings.</b> Accept (%g+): usable as received. '
               'Minor fixes (%g&ndash;%g): usable; the GP should correct it next time. '
               'Resubmit (%g&ndash;%g): not usable until corrected. Reject (below %g): wrong or '
               'unreadable file. Usable = Accept + Minor fixes. Right first time = Accept.<br>'
               'File-level detail: summary.xlsx and run_summary.html in the same folder. '
               'Input folder: %s</div>'
               % (cfg.bands['accept'], cfg.bands['minor_fixes'], cfg.bands['accept'] - 1,
                  cfg.bands['resubmit'], cfg.bands['minor_fixes'] - 1, cfg.bands['resubmit'],
                  html.escape(str(run_info.get('folder', '')))))
    out.append('</body></html>')

    with open(path, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(out))
    return path
