"""The short recap printed when a run ends, plus a CSV of the summary table."""
from __future__ import annotations

import collections
import csv
from typing import List

from ..scorer import FileScore

SUMMARY_COLUMNS = ['file', 'score', 'rating', 'version', 'positions', 'columns_mapped',
                   'structure', 'completeness', 'validity', 'consistency',
                   'gates', 'not_scored']


def write_csv(path: str, scores: List[FileScore]) -> str:
    with open(path, 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.writer(fh)
        w.writerow(SUMMARY_COLUMNS)
        for s in sorted(scores, key=lambda x: (-x.score, x.filename)):
            w.writerow(s.as_summary_row())
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
