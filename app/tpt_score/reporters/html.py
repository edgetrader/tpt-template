"""One-page, self-contained HTML summary. Print it to PDF for the one-pager."""
from __future__ import annotations

import collections
import html
from typing import List

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
    worst = sorted(scores, key=lambda x: (x.score, x.filename))[:10]
    if worst:
        out.append('<section><h2>Files needing attention, lowest first</h2><table>'
                   '<colgroup><col style="width:36%"><col style="width:62px">'
                   '<col style="width:96px"><col></colgroup>'
                   '<tr><th>File</th><th class="n">Score</th><th>Rating</th>'
                   '<th>Main reason</th></tr>')
        for s in worst:
            reason = s.error or (s.gates[0] if s.gates else
                                 (s.top_fixes[0].detail or s.top_fixes[0].title
                                  if s.top_fixes else 'no check failed'))
            out.append('<tr><td class="clip">%s</td><td class="n">%.1f</td><td>%s</td>'
                       '<td class="clip">%s</td></tr>'
                       % (html.escape(s.filename), s.score, _pill(s.rating),
                          html.escape(str(reason))))
        out.append('</table></section>')

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
