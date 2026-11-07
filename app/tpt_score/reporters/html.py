"""One-page, self-contained HTML summary. Print it to PDF for the one-pager."""
from __future__ import annotations

import collections
import html
import os
from typing import List

from ..remarks import (FINDING_LABELS, REMARK_LABELS, column_findings, dates_standard,
                       gate_reference, remark_findings, template_problems)
from .console import usable_files
from ..scorer import DIMENSIONS, DIMENSION_LABEL, FileScore

BAND_COLOUR = {'Accept': '#2e7d32', 'Minor fixes': '#b58100',
               'Resubmit': '#d4760a', 'Reject': '#c62828'}

CSS = """
*{box-sizing:border-box}
body{font:13px/1.45 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;color:#1a1a1a;
     margin:0;padding:22px 26px;background:#fff}
h1{font-size:19px;margin:0 0 2px}
.sub{color:#666;font-size:11.5px;margin-bottom:14px}
.tiles{display:grid;grid-template-columns:repeat(6,1fr);gap:9px;margin-bottom:16px}
.tile{border:1px solid #d8d8d8;border-radius:8px;padding:8px 11px}
.tile .n{font-size:21px;font-weight:600;line-height:1.15}
.tile .l{font-size:10px;color:#666;text-transform:uppercase;letter-spacing:.04em;
         white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tile.band{border-left-width:4px}
h2{font-size:12.5px;margin:0 0 6px;padding-bottom:3px;border-bottom:1px solid #e3e3e3}
section{margin-bottom:15px}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:24px;align-items:start}
table{border-collapse:collapse;width:100%;font-size:11.5px;table-layout:fixed}
th{text-align:left;color:#666;font-weight:600;padding:3px 8px 3px 0;border-bottom:1px solid #e3e3e3;
   white-space:nowrap}
td{padding:3px 8px 3px 0;border-bottom:1px solid #f1f1f1;vertical-align:middle}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap;padding-right:12px}
td.clip{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
td.wrap{white-space:normal;word-break:break-word;vertical-align:top;line-height:1.35}
td.top{vertical-align:top}
.gate{font-weight:600;white-space:nowrap}
.tab{display:inline-block;background:#ddebf7;color:#1f3864;font-weight:600;border-radius:3px;padding:0 5px}
.f{margin:0 0 3px}.f b{font-weight:600}
.f.missing b{background:#fcd5b4}.f.extra b{background:#fff2cc}
.f.renamed b{background:#ffeb9c}.f.blanks b{background:#ffc7CE}
.f.types b{background:#f4b6c2}.f.dates b{background:#ffe699}.f.dates_ok b{background:#e2efda}
.f.bad b{background:#ffc7ce;color:#9c0006}.f.layout b{background:#e7e6e6}.f.checks b{background:#dde5f0}
.f b{padding:0 4px;border-radius:2px}
details{margin-top:4px}summary{cursor:pointer;color:#1f3864;font-weight:600;font-size:11.5px}
.paths{font:10.5px/1.5 ui-monospace,Menlo,Consolas,monospace;word-break:break-all;margin-top:5px}
.track{background:#f1f1f1;border-radius:3px;height:14px;width:100%}
.bar{height:14px;border-radius:3px;display:block}
.pill{display:inline-block;padding:1px 7px;border-radius:9px;font-size:10.5px;color:#fff;
      white-space:nowrap}
.foot{margin-top:16px;color:#888;font-size:10.5px;border-top:1px solid #e3e3e3;padding-top:7px;
      word-break:break-all}
@media print{body{padding:0;font-size:11.5px}h1{font-size:17px}}
"""


def _pill(rating: str) -> str:
    return '<span class="pill" style="background:%s">%s</span>' % (
        BAND_COLOUR.get(rating, '#777'), html.escape(rating))


def _bar(pct: float, colour: str) -> str:
    """A full-width track with the bar filling its share of it."""
    return ('<div class="track"><span class="bar" style="width:%.1f%%;background:%s"></span></div>'
            % (max(0.0, min(100.0, pct)), colour))


# Highlight class per finding, so every report colours a finding the same way.
LINE_CLASS = {'notscored': 'bad', 'template': 'bad', 'tables': 'bad', 'layout': 'layout',
              'repeated': 'extra', 'missing': 'missing', 'extra': 'extra', 'renamed': 'renamed',
              'blanks': 'blanks', 'types': 'types', 'checks': 'checks'}
ORDER = ['notscored', 'template', 'tables', 'missing', 'extra', 'renamed', 'repeated', 'blanks',
         'types', 'dates', 'layout', 'checks']


def _line(cls: str, label: str, items, limit: int = 6) -> str:
    """<b>Label (n)</b> item; item; ... and N more: the one format for every finding."""
    more = ' and %d more' % (len(items) - limit) if len(items) > limit else ''
    return ('<div class="f %s"><b>%s (%d)</b> %s%s</div>'
            % (cls, html.escape(label), len(items), html.escape('; '.join(items[:limit])), more))


def finding_lines(s) -> str:
    """Every finding for one file, in a fixed order and one format."""
    cols = column_findings(s)
    rems = remark_findings(s)
    labels = dict(FINDING_LABELS)
    labels.update(dict(REMARK_LABELS))
    cell = []
    head = []
    if s.fund_ref_status:
        if s.fund_ref_status in ('exact', 'close'):
            head.append('GP: <b>%s</b>' % html.escape(s.gp_name))
        else:
            head.append('GP: <b class="down">not identified</b>')
    if s.sheet:
        head.append('Tab: <span class="tab">%s</span>' % html.escape(s.sheet))
    if head:
        cell.append('<div class="f">%s</div>' % ' &middot; '.join(head))
    if s.fund_ref_status:
        text = {'exact': 'matched to %s' % s.ref_fund,
                'close': 'close match to %s, please confirm' % s.ref_fund,
                'ambiguous': 'matches more than one GP: %s' % s.fund_ref_label.split(': ', 1)[-1],
                'no name': 'no fund name in the file name'}.get(
            s.fund_ref_status, 'fund in the file name is not in the fund reference')
        cls = 'dates_ok' if s.fund_ref_status == 'exact' else (
            'renamed' if s.fund_ref_status == 'close' else 'bad')
        cell.append('<div class="f %s"><b>Fund reference</b> %s</div>' % (cls, html.escape(text)))
    for key in ORDER:
        if key == 'dates':
            items = cols['dates']
            if items:
                if dates_standard(s):
                    cell.append('<div class="f dates_ok"><b>Date formats (%d)</b> all yyyy-mm-dd</div>'
                                % len(items))
                else:
                    off = [i for i, c in zip(items, s.date_formats) if not c[4]]
                    cell.append(_line('dates', 'Date formats not yyyy-mm-dd', off, 4))
            continue
        items = cols.get(key) if key in cols else rems.get(key, [])
        if items:
            cell.append(_line(LINE_CLASS[key], labels[key], items))
    if not any(cols[k] for k in cols if k != 'dates') and not any(rems.values()):
        cell.append('<div class="f">No issues found.</div>')
    return ''.join(cell)


def write_html(path: str, scores: List[FileScore], cfg, run_info: dict) -> str:
    n = len(scores)
    vals = [s.score for s in scores]
    avg = sum(vals) / n if n else 0.0
    counts = collections.Counter(s.rating for s in scores)
    unreadable = [s for s in scores if s.error]
    scored = [s for s in scores if not s.error and s.dimensions]

    issues = collections.Counter()
    titles = {}
    for s in scores:
        for cid, cs in s.checks.items():
            if cs.f > 0:
                issues[cid] += 1
                titles[cid] = cs.title

    out = ['<!DOCTYPE html><html><head><meta charset="utf-8">',
           '<title>TPT submission quality: run summary</title>',
           '<style>%s</style></head><body>' % CSS,
           '<h1>TPT submission quality: run summary</h1>',
           '<div class="sub">%s &middot; %d files &middot; rules version %s &middot; '
           '%d could not be scored</div>' % (
               html.escape(str(run_info.get('started', ''))), n,
               html.escape(cfg.rules_version), len(unreadable)),
           '<div class="tiles">']
    out.append('<div class="tile"><div class="n">%d</div><div class="l">files scored</div></div>' % n)
    out.append('<div class="tile"><div class="n">%.1f</div><div class="l">average score</div></div>' % avg)
    for rating in ('Accept', 'Minor fixes', 'Resubmit', 'Reject'):
        c = counts.get(rating, 0)
        out.append('<div class="tile band" style="border-left-color:%s">'
                   '<div class="n">%d</div><div class="l">%s</div></div>'
                   % (BAND_COLOUR[rating], c, html.escape(rating.lower())))
    out.append('</div>')

    # ---- two columns: how the run landed, and what went wrong most often ----
    out.append('<div class="cols"><div>')

    out.append('<section><h2>Score distribution</h2><table>'
               '<colgroup><col style="width:86px"><col style="width:34px">'
               '<col style="width:40px"><col></colgroup>')
    for rating in ('Accept', 'Minor fixes', 'Resubmit', 'Reject'):
        c = counts.get(rating, 0)
        pct = 100.0 * c / max(1, n)
        out.append('<tr><td>%s</td><td class="n">%d</td><td class="n">%.0f%%</td><td>%s</td></tr>'
                   % (_pill(rating), c, pct, _bar(pct, BAND_COLOUR[rating])))
    out.append('</table></section>')

    if scored:
        # Dimension averages cluster near 100, so a bar drawn from zero shows no
        # difference between them. The bar is the shortfall instead, scaled so
        # the weakest dimension fills the track: it answers "which is worst".
        avgs = {d: sum(s.dimensions[d].score for s in scored) / len(scored)
                for d in DIMENSIONS}
        worst_gap = max(100.0 - a for a in avgs.values()) or 1.0
        out.append('<section><h2>Average by dimension, across the %d files that could be '
                   'scored</h2><table>'
                   '<colgroup><col style="width:124px"><col style="width:52px"><col></colgroup>'
                   % len(scored))
        for d in DIMENSIONS:
            a = avgs[d]
            colour = ('#2e7d32' if a >= 95 else '#b58100' if a >= 80 else '#c62828')
            out.append('<tr><td>%s</td><td class="n">%.1f</td><td>%s</td></tr>'
                       % (html.escape(DIMENSION_LABEL[d]), a,
                          _bar((100.0 - a) / worst_gap * 100.0, colour)))
        out.append('</table>'
                   '<div style="color:#888;font-size:10.5px;margin-top:3px">'
                   'Bar shows the shortfall from 100, the weakest dimension filling the '
                   'track. Shorter is better.</div></section>')

    out.append('</div><div>')
    out.append('<section><h2>Most frequent issues</h2><table>'
               '<colgroup><col style="width:46px"><col><col style="width:44px"></colgroup>'
               '<tr><th>Check</th><th>Name</th><th class="n">Files</th></tr>')
    for cid, c in issues.most_common(12):
        out.append('<tr><td>%s</td><td class="clip">%s</td><td class="n">%d</td></tr>'
                   % (cid, html.escape(titles.get(cid, '')), c))
    if not issues:
        out.append('<tr><td colspan="3">No check failed in this run.</td></tr>')
    out.append('</table></section></div></div>')

    # ---- full width, where the long text has room -------------------------
    need = sorted((x for x in scores if x.rating != 'Accept'), key=lambda x: (x.score, x.filename))
    shown = need[:30]
    if shown:
        out.append('<section><h2>Files needing attention, lowest first</h2><table>'
                   '<colgroup><col style="width:24%"><col style="width:54px">'
                   '<col style="width:92px"><col></colgroup>'
                   '<tr><th>File</th><th class="n">Score</th><th>Rating</th>'
                   '<th>Findings against the V7 template</th></tr>')
        for s in shown:
            out_cell = finding_lines(s)
            out.append('<tr><td class="wrap top">%s</td><td class="n top">%.1f</td>'
                       '<td class="top">%s</td><td class="wrap">%s</td></tr>'
                       % (html.escape(s.filename), s.score, _pill(s.rating), out_cell))
        out.append('</table>')
        if len(need) > len(shown):
            out.append('<div style="color:#888;font-size:10.5px;margin-top:3px">%d more files '
                       'below Accept are listed in summary.xlsx.</div>' % (len(need) - len(shown)))
        out.append('</section>')

    usable = usable_files(scores)
    out.append('<section><h2>Files rated Accept or Minor fixes: %d of %d</h2>' % (len(usable), n))
    out.append('<details><summary>Show the full path of each file (also in accepted_files.xlsx '
               'and the Accepted files sheet)</summary><div class="paths">%s</div></details></section>'
               % '<br>'.join(html.escape(os.path.abspath(u.path)) for u in usable))

    out.append('<section><h2>Gates: hard rules applied on top of the score</h2>'
               '<table><colgroup><col style="width:44px"><col style="width:36%">'
               '<col style="width:110px"><col></colgroup>'
               '<tr><th>Gate</th><th>What triggers it</th><th>Effect</th><th>In plain words</th></tr>')
    for gate, trigger, effect, plain in gate_reference(cfg):
        out.append('<tr><td class="gate top">%s</td><td class="wrap">%s</td>'
                   '<td class="wrap">%s</td><td class="wrap">%s</td></tr>'
                   % (gate, html.escape(trigger), html.escape(effect), html.escape(plain)))
    out.append('</table><div style="color:#888;font-size:10.5px;margin-top:3px">'
               'A gate sets the score to 0 or caps it, whatever the weighted score says. '
               'The lowest cap that applies wins.</div></section>')

    out.append('<div class="foot">Folder: %s &middot; template: %s &middot; config: %s<br>'
               'Bands: Accept %g+, Minor fixes %g to %g, Resubmit %g to %g, Reject below %g. '
               'A score of 0 means the file could not be used at all.</div>'
               % (html.escape(str(run_info.get('folder', ''))),
                  html.escape(str(run_info.get('template', ''))),
                  html.escape(str(run_info.get('config', ''))),
                  cfg.bands['accept'], cfg.bands['minor_fixes'], cfg.bands['accept'] - 1,
                  cfg.bands['resubmit'], cfg.bands['minor_fixes'] - 1, cfg.bands['resubmit']))
    out.append('</body></html>')

    with open(path, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(out))
    return path
