"""The Excel scorecard workbook."""
from __future__ import annotations

import os
from typing import List

import openpyxl
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as CL

from ..checks import CHECKS
from ..scorer import DIMENSIONS, DIMENSION_LABEL, FileScore

HEAD_FILL = PatternFill('solid', fgColor='1F3864')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=10)
TITLE_FONT = Font(bold=True, size=13)
SUB_FONT = Font(italic=True, color='595959', size=9)
BOLD = Font(bold=True)

RATING_FILL = {
    'Accept': PatternFill('solid', fgColor='C6EFCE'),
    'Minor fixes': PatternFill('solid', fgColor='FFEB9C'),
    'Resubmit': PatternFill('solid', fgColor='FCD5B4'),
    'Reject': PatternFill('solid', fgColor='FFC7CE'),
}
STATUS_FILL = {
    'Pass': PatternFill('solid', fgColor='E2EFDA'),
    'Minor': PatternFill('solid', fgColor='FFF2CC'),
    'Fail': PatternFill('solid', fgColor='FCE4E4'),
    'Not evaluated': PatternFill('solid', fgColor='F2F2F2'),
}


def _header(ws, row, labels, widths=None):
    for i, text in enumerate(labels, start=1):
        c = ws.cell(row, i, text)
        c.fill, c.font = HEAD_FILL, HEAD_FONT
        c.alignment = Alignment(vertical='center', wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[CL(i)].width = w
    ws.freeze_panes = ws.cell(row + 1, 1)


def _safe_sheet_name(name: str, used) -> str:
    base = ''.join(ch for ch in name if ch not in '[]:*?/\\').strip()[:28]
    if not base:
        base = 'file'                 # Excel refuses an empty sheet name
    out, n = base, 1
    while out.lower() in used:
        out = '%s~%d' % (base[:26], n)
        n += 1
    used.add(out.lower())
    return out


def write_workbook(path: str, scores: List[FileScore], cfg, run_info: dict):
    wb = openpyxl.Workbook()
    used = set()

    _run_summary(wb.active, scores, cfg, run_info)
    used.add(wb.active.title.lower())
    _dimension_scores(wb.create_sheet('Dimension scores'), scores)
    used.add('dimension scores')
    _check_matrix(wb.create_sheet('Check matrix'), scores)
    used.add('check matrix')
    _issue_log(wb.create_sheet('Issue log'), scores)
    used.add('issue log')
    _feedback(wb.create_sheet('Feedback'), scores)
    used.add('feedback')

    limit = int(cfg.data.get('reports', {}).get('file_scorecard_sheets_max', 150))
    for fscore in sorted(scores, key=lambda s: s.score)[:limit]:
        name = _safe_sheet_name(os.path.splitext(fscore.filename)[0], used)
        _scorecard(wb.create_sheet(name), fscore, cfg)

    wb.save(path)
    return path


# --------------------------------------------------------------------------- sheets
def _run_summary(ws, scores, cfg, run_info):
    ws.title = 'Run summary'
    ws['A1'] = 'TPT submission quality: run summary'
    ws['A1'].font = TITLE_FONT
    meta = [
        ('Run date', run_info.get('started', '')),
        ('Input folder', run_info.get('folder', '')),
        ('Files scored', len(scores)),
        ('Rules version', cfg.rules_version),
        ('Config', run_info.get('config', '')),
        ('Template', run_info.get('template', '')),
    ]
    usable = [s.score for s in scores]
    if usable:
        ordered = sorted(usable)
        meta += [
            ('Average score', round(sum(usable) / len(usable), 1)),
            ('Median score', round(ordered[len(ordered) // 2], 1)),
            ('Lowest score', round(ordered[0], 1)),
        ]
    r = 3
    for k, v in meta:
        ws.cell(r, 1, k).font = BOLD
        ws.cell(r, 2, v)
        r += 1

    r += 1
    ws.cell(r, 1, 'Files per rating').font = BOLD
    r += 1
    for rating in ('Accept', 'Minor fixes', 'Resubmit', 'Reject'):
        n = sum(1 for s in scores if s.rating == rating)
        ws.cell(r, 1, rating).fill = RATING_FILL[rating]
        ws.cell(r, 2, n)
        ws.cell(r, 3, '%.0f%%' % (100.0 * n / max(1, len(scores))))
        r += 1

    r += 1
    start = r
    _header(ws, r, ['File', 'Score', 'Rating', 'Version', 'Positions', 'Columns mapped',
                    'Structure', 'Completeness', 'Validity', 'Consistency',
                    'Gates applied', 'Could not be read'],
            [46, 8, 13, 9, 10, 14, 11, 13, 10, 12, 44, 40])
    ws.freeze_panes = ws.cell(start + 1, 1)
    r += 1
    for s in sorted(scores, key=lambda x: (-x.score, x.filename)):
        for i, v in enumerate(s.as_summary_row(), start=1):
            ws.cell(r, i, v)
        ws.cell(r, 3).fill = RATING_FILL.get(s.rating, STATUS_FILL['Not evaluated'])
        ws.cell(r, 2).number_format = '0.0'
        r += 1
    if r > start + 1:
        ws.conditional_formatting.add('B%d:B%d' % (start + 1, r - 1),
                                      ColorScaleRule(start_type='num', start_value=0, start_color='F8696B',
                                                     mid_type='num', mid_value=80, mid_color='FFEB84',
                                                     end_type='num', end_value=100, end_color='63BE7B'))
    ws.auto_filter.ref = 'A%d:L%d' % (start, max(start, r - 1))


def _dimension_scores(ws, scores):
    _header(ws, 1, ['File', 'Structure (15)', 'Completeness (30)', 'Validity (30)',
                    'Consistency (25)', 'Total', 'Rating', 'Cap applied'],
            [46, 14, 17, 13, 15, 9, 13, 12])
    r = 2
    for s in sorted(scores, key=lambda x: (-x.score, x.filename)):
        ws.cell(r, 1, s.filename)
        for i, d in enumerate(DIMENSIONS, start=2):
            v = round(s.dimensions[d].score, 1) if s.dimensions else ''
            ws.cell(r, i, v).number_format = '0.0'
        ws.cell(r, 6, round(s.score, 1)).number_format = '0.0'
        ws.cell(r, 7, s.rating).fill = RATING_FILL.get(s.rating, STATUS_FILL['Not evaluated'])
        ws.cell(r, 8, s.cap if s.cap is not None else '')
        r += 1
    if r > 2:
        ws.conditional_formatting.add('B2:F%d' % (r - 1),
                                      ColorScaleRule(start_type='num', start_value=0, start_color='F8696B',
                                                     mid_type='num', mid_value=80, mid_color='FFEB84',
                                                     end_type='num', end_value=100, end_color='63BE7B'))
    ws.auto_filter.ref = 'A1:H%d' % max(1, r - 1)


def _check_matrix(ws, scores):
    ids = sorted(CHECKS, key=lambda k: (k[0], int(k[1:])))
    ws.cell(1, 1, 'Share of applicable rows or cells failing each check (0 = clean, 1 = all)').font = SUB_FONT
    _header(ws, 2, ['File'] + ids, [46] + [7] * len(ids))
    for i, cid in enumerate(ids, start=2):
        ws.cell(2, i).alignment = Alignment(horizontal='center')
    r = 3
    for s in sorted(scores, key=lambda x: (-x.score, x.filename)):
        ws.cell(r, 1, s.filename)
        for i, cid in enumerate(ids, start=2):
            cs = s.checks.get(cid)
            if cs is None:
                ws.cell(r, i, '').fill = STATUS_FILL['Not evaluated']
            else:
                c = ws.cell(r, i, round(cs.f, 3))
                c.number_format = '0.000'
                c.fill = STATUS_FILL.get(cs.status, STATUS_FILL['Not evaluated'])
        r += 1
    if r > 3:
        ws.conditional_formatting.add('B3:%s%d' % (CL(len(ids) + 1), r - 1),
                                      ColorScaleRule(start_type='num', start_value=0, start_color='FFFFFF',
                                                     end_type='num', end_value=1, end_color='F8696B'))
    ws.freeze_panes = 'B3'


def _issue_log(ws, scores):
    _header(ws, 1, ['File', 'Check', 'Check name', 'Severity', 'Sheet row', 'Column',
                    'Value found', 'Expected'],
            [40, 7, 40, 10, 10, 34, 34, 44])
    r = 2
    for s in sorted(scores, key=lambda x: x.filename):
        if s.error and not s.checks:
            ws.cell(r, 1, s.filename)
            ws.cell(r, 2, s.gates[0].split(':')[0] if s.gates else '')
            ws.cell(r, 3, 'File could not be scored')
            ws.cell(r, 4, 'Fail')
            ws.cell(r, 8, s.error)
            ws.cell(r, 4).fill = STATUS_FILL['Fail']
            r += 1
            continue
        for cid in sorted(s.checks, key=lambda k: (k[0], int(k[1:]))):
            cs = s.checks[cid]
            for issue in cs.issues:
                ws.cell(r, 1, s.filename)
                ws.cell(r, 2, cid)
                ws.cell(r, 3, cs.title)
                ws.cell(r, 4, cs.status).fill = STATUS_FILL.get(cs.status, STATUS_FILL['Not evaluated'])
                for i, v in enumerate(issue.as_row(), start=5):
                    ws.cell(r, i, v)
                r += 1
                if r > 50000:
                    ws.cell(r, 1, '(issue log truncated at 50,000 rows)')
                    return
    ws.auto_filter.ref = 'A1:H%d' % max(1, r - 1)


def _feedback(ws, scores):
    _header(ws, 1, ['File', 'Fund', 'Score', 'Rating', 'What to fix, in order of impact'],
            [40, 32, 8, 13, 110])
    r = 2
    for s in sorted(scores, key=lambda x: x.score):
        ws.cell(r, 1, s.filename)
        ws.cell(r, 2, s.fund_name)
        ws.cell(r, 3, round(s.score, 1)).number_format = '0.0'
        ws.cell(r, 4, s.rating).fill = RATING_FILL.get(s.rating, STATUS_FILL['Not evaluated'])
        ws.cell(r, 5, feedback_text(s)).alignment = Alignment(wrap_text=True, vertical='top')
        ws.row_dimensions[r].height = 76
        r += 1


def feedback_text(s: FileScore) -> str:
    if s.error and not s.checks:
        return 'The file could not be scored: %s' % s.error
    lines = []
    for i, c in enumerate(s.top_fixes, start=1):
        lines.append('%d. %s - %s (costs %.1f points)' % (i, c.title, c.detail, c.points_lost))
    if not lines:
        lines.append('No issues found.')
    return '\n'.join(lines)


def _scorecard(ws, s: FileScore, cfg):
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 46
    for col, w in zip('CDEFG', (11, 11, 11, 13, 54)):
        ws.column_dimensions[col].width = w

    ws['A1'] = 'Scorecard: %s' % s.filename
    ws['A1'].font = TITLE_FONT
    ws['A2'] = 'Rules version %s, scored %s' % (s.rules_version, s.scored_at)
    ws['A2'].font = SUB_FONT

    r = 4
    for k, v in [('Fund', s.fund_name), ('Fund identifier', s.fund_id),
                 ('Template version', s.version), ('Sheet', s.sheet),
                 ('Header cell', s.header_cell), ('Positions read', s.positions),
                 ('Columns mapped', s.mapped_columns)]:
        ws.cell(r, 1, k).font = BOLD
        ws.cell(r, 2, v)
        r += 1

    r += 1
    ws.cell(r, 1, 'Score').font = BOLD
    c = ws.cell(r, 2, round(s.score, 1))
    c.font = Font(bold=True, size=18)
    c.number_format = '0.0'
    ws.cell(r, 3, s.rating).fill = RATING_FILL.get(s.rating, STATUS_FILL['Not evaluated'])
    ws.cell(r, 3).font = BOLD
    r += 1
    if s.gates:
        ws.cell(r, 1, 'Gates').font = BOLD
        ws.cell(r, 2, '; '.join(s.gates))
        r += 1
    if s.error:
        ws.cell(r, 1, 'Not scored').font = BOLD
        ws.cell(r, 2, s.error)
        r += 1

    r += 1
    _header(ws, r, ['', 'Dimension', 'Weight', 'Score', 'Contribution', '', ''])
    r += 1
    for d in DIMENSIONS:
        ds = s.dimensions.get(d)
        ws.cell(r, 2, DIMENSION_LABEL[d])
        if ds:
            ws.cell(r, 3, ds.weight)
            ws.cell(r, 4, round(ds.score, 1)).number_format = '0.0'
            ws.cell(r, 5, round(ds.contribution, 2)).number_format = '0.00'
        r += 1
    ws.cell(r, 2, 'Total').font = BOLD
    ws.cell(r, 5, round(s.score, 1)).number_format = '0.0'
    ws.cell(r, 5).font = BOLD
    r += 2

    if s.checks:
        _header(ws, r, ['Check', 'Name', 'Share failed', 'Points lost', 'Affected',
                        'Status', 'Detail'])
        r += 1
        for cid in sorted(s.checks, key=lambda k: (k[0], int(k[1:]))):
            cs = s.checks[cid]
            ws.cell(r, 1, cid + (' *' if cs.critical else ''))
            ws.cell(r, 2, cs.title)
            ws.cell(r, 3, round(cs.f, 3)).number_format = '0.000'
            ws.cell(r, 4, round(cs.points_lost, 2)).number_format = '0.00'
            ws.cell(r, 5, '%d of %d' % (cs.n_affected, cs.n_applicable) if cs.n_applicable else '')
            ws.cell(r, 6, cs.status).fill = STATUS_FILL.get(cs.status, STATUS_FILL['Not evaluated'])
            ws.cell(r, 7, cs.detail)
            r += 1
        ws.cell(r, 1, '* critical check: any failure caps the score and cannot reach Accept').font = SUB_FONT
        r += 2

    ws.cell(r, 1, 'Top fixes').font = BOLD
    r += 1
    for line in feedback_text(s).split('\n'):
        ws.cell(r, 1, line)
        r += 1

    r += 1
    ws.cell(r, 1, 'Recovery log').font = BOLD
    r += 1
    for note in (s.recovery or ['nothing had to be repaired']):
        ws.cell(r, 1, note)
        r += 1
