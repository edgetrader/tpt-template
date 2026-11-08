"""The short recap printed when a run ends, plus a CSV of the summary table."""
from __future__ import annotations

import collections
import csv
import os
from typing import List

from ..scorer import FileScore

SUMMARY_COLUMNS = ['file', 'gp_name', 'fund_from_file_name', 'fund_reference_match', 'score', 'rating', 'version', 'positions', 'columns_mapped',
                   'structure', 'completeness', 'validity', 'consistency',
                   'gates', 'not_scored', 'tab_found_in', 'key_fields', 'missing_columns', 'extra_columns',
                   'renamed_columns', 'mandatory_blanks', 'wrong_data_type', 'date_formats',
                   'remarks']


def write_csv(path: str, scores: List[FileScore]) -> str:
    with open(path, 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.writer(fh)
        w.writerow(SUMMARY_COLUMNS)
        for s in sorted(scores, key=lambda x: (-x.score, x.filename)):
            w.writerow(s.as_summary_row())
    return path


USABLE_RATINGS = ('Accept', 'Minor fixes')


def usable_files(scores: List[FileScore]) -> List[FileScore]:
    """Files rated Accept or Minor fixes, in path order."""
    return sorted((s for s in scores if s.rating in USABLE_RATINGS), key=lambda x: os.path.abspath(x.path))


def write_usable_list(path: str, scores: List[FileScore]) -> str:
    """One full path per line: the files that can go forward."""
    with open(path, 'w', encoding='utf-8') as fh:
        for s in usable_files(scores):
            fh.write(os.path.abspath(s.path) + '\n')
    return path


def recap(scores: List[FileScore], cfg, outputs: dict) -> str:
    n = len(scores)
    vals = [s.score for s in scores]
    counts = collections.Counter(s.rating for s in scores)
    lines = [
        '%d files scored, average %.1f' % (n, (sum(vals) / n) if n else 0.0),
        '  ' + ', '.join('%s %d' % (r, counts.get(r, 0))
                         for r in ('Accept', 'Minor fixes', 'Resubmit', 'Reject')),
    ]
    worst = sorted(scores, key=lambda x: x.score)[:3]
    if worst:
        lines.append('  lowest: ' + ', '.join('%s %.1f' % (s.filename, s.score) for s in worst))
    for label, path in outputs.items():
        if path:
            lines.append('  %-9s %s' % (label + ':', path))
    return '\n'.join(lines)
