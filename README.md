# TPT Template: data quality checking and scoring

Checks and scores the data files General Partners submit on the FinDatEx
Tripartite Template (TPT) V7. Point the tool at a folder of submissions; every
file comes back with a score from 0 to 100, a rating, and a list of what the GP
needs to fix.

```
cd app
python tpt_score.py --in ../data
```

Reports land in `output/<input folder>/run_<timestamp>/`, e.g. `output/data/run_20261107_221500/`.

## What is here

| Folder | What it holds |
| --- | --- |
| `app/` | The scoring system: scripts, `rules.yaml`, tests and fixtures. See `app/README.md` |
| `data/` | TPT files to score. Currently 100 synthetic V7 deliveries with known faults, plus their ground truth. See `data/README.md` |
| `output/` | Where each run writes its reports: one folder per input folder, one timestamped subfolder per run |
| `templates/` | The FinDatEx workbooks, V5 to V8. V7 is the one the rules are read from |
| `archive/` | Earlier sample sets and the original checker. Not used by the scorer |
| `CLAUDE.md` | The design: scoring model, the 24 checks, gates, build plan |

## Running it

```
cd app
python tpt_score.py --in <folder>              # score every file in a folder
python tpt_score.py --in <folder> --recursive  # include subfolders
python tpt_score.py --in <one file.xlsx>       # score a single file
python tpt_score.py --help                     # every option
```

The V7 workbook is found automatically in `templates/`. Needs Python 3.9 or
later with `openpyxl` and `pyyaml`:

```
pip install openpyxl pyyaml
```

Nothing leaves the machine, and no submission file is ever modified.

Useful options:

| Option | What it does |
| --- | --- |
| `-r`, `--recursive` | Also score files in subfolders |
| `--pattern GLOB` | Only files matching a name pattern, e.g. `--pattern 'TPT_V7_SYN_*.xlsx'` |
| `-o`, `--out FOLDER` | Send the reports somewhere other than `output/` |
| `-c`, `--config FILE` | Use a different `rules.yaml` |
| `-q`, `--quiet` | Print only the closing recap |

Exit code 0 when every file is Accept or Minor fixes, 1 when something needs
action, 2 on a path or configuration error, so a run can be chained.

## What a run produces

| File | What it is |
| --- | --- |
| `TPT_scorecards.xlsx` | Run summary, dimension scores, check matrix, issue log, GP feedback, and one scorecard sheet per file |
| `run_summary.html` | One page for management. Self-contained; print it to PDF |
| `summary.csv` | The summary table, for loading elsewhere |

## How the score works

Four weighted dimensions, each starting at 100 and losing points per failed
check:

| Dimension | Weight | What it measures |
| --- | --- | --- |
| Structure | 15 | Whether the table can be found and mapped to the V7 template |
| Completeness | 30 | Whether required values are present for every position |
| Validity and format | 30 | Whether values have the right type, format and code |
| Consistency | 25 | Whether values agree with each other and with the fund totals |

Most checks charge the **share** of rows or cells that failed, so a large file
does not score worse than a small one of the same quality. Six gates sit on top:
an unreadable file, no recognisable header, or no positions scores 0; a
non-V7 template caps at 40; losing more than a fifth of the mandatory columns
caps at 49; and any critical check failing caps at 94, so Accept is out of reach.

**95+ Accept, 80 to 94 Minor fixes, 50 to 79 Resubmit, below 50 Reject.**

A file whose table does not start in A1 is still read: the loader finds the
header, skips title rows, totals rows and notes, and maps columns by field
number, template name, GP short label, then fuzzy match. What it repaired is
written to the recovery log on that file's scorecard and charged under
Structure, so the content is still checked.

Every weight, tolerance and threshold lives in `app/rules.yaml`, and the rules
version is printed in every report, so any score can be traced to the settings
that produced it. `CLAUDE.md` lists all 24 checks and their weights.

## Tests

```
cd app
python tests/test_tpt_score.py
```

65 tests: 33 fixtures (one clean file that scores exactly 100.0, one per check,
one per gate), the value parsers against real ISINs and LEIs, the scoring
arithmetic, monotonicity, repeatability, an end-to-end run, a calibration pass
over `data/`, and a regression test for each defect found in review.

## Status

Built and tested. Weights, gate caps and tolerances are starting values and
have not yet been calibrated against real GP submissions — the open points are
listed at the end of `CLAUDE.md`.
