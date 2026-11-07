# tpt_score

Checks and scores the data files General Partners submit on the FinDatEx
Tripartite Template (TPT) V7. Point it at a folder; every file comes back with
a score from 0 to 100, a rating, and a list of what to fix.

```
cd app
python tpt_score.py --in ../data
```

The scripts live in `app/`; everything a run produces goes to `output/` beside
it, in a timestamped folder holding:

| File | What it is |
| --- | --- |
| `management_summary.html` | One page for management: usable %, right first time, average score and files to be resent, each with the change since the last run over the same folder; problems in business terms; results by fund (or by GP with a `gp_map.csv`); the files to be resent with what to ask the GP |
| `run_stats.json` | The run's figures, read by the next run for the comparison |
| `TPT_scorecards.xlsx` | Run summary, gate reference, dimension scores, check matrix, issue log, feedback, and one scorecard sheet per file |
| `run_summary.html` | One-page summary for management, with a gate reference. Print it to PDF |
| `summary.xlsx` | One sheet: the summary table with a Remarks column, and the gate codes explained above it |
| `accepted_files.txt` | Full path of every file rated Accept or Minor fixes, one per line (also a sheet in the workbooks and a list in the HTML) |
| `summary.csv` | Same table, only when you pass `--csv` |

The summary table lists, per file: the **tab** the template was found in, **missing columns** (mandatory ones flagged), **extra columns** not in V7, **renamed columns** (found -> template name), **mandatory columns with blanks**, **wrong data type** (columns whose cells do not match the expected type, with an example), the **format of every date column**, and a short **remarks** line. A different template or more than one TPT table is shown in capitals and highlighted red. The full lists are on the *Column findings* sheet.

## Installing

Python 3.9 or later, plus two packages:

```
pip install openpyxl pyyaml
```

Nothing leaves the machine and no submission file is ever modified.

## Options

```
python tpt_score.py --in FOLDER [options]

  -i, --in FOLDER        folder holding the TPT files (a single file also works)
  -o, --out FOLDER       where the reports go (default: the output folder
                         beside the app folder)
  -t, --template FILE    the FinDatEx V7 workbook the rules are read from
  -c, --config FILE      rules.yaml with the weights and thresholds
  -r, --recursive        also score files in subfolders
      --pattern GLOB     only files matching this pattern
      --no-excel         skip the Excel workbook
      --no-html          skip the HTML summary
      --no-timestamp     write straight into --out
  -q, --quiet            only print the recap
```

The exit code is 0 when every file is Accept or Minor fixes, 1 when at least
one needs action, and 2 on a configuration or path error, so a run can be
chained into a pipeline.

The V7 workbook is found automatically when it sits in `templates/` next to
the submissions folder. Otherwise pass `--template`.


### Grouping the management page by GP

Copy `gp_map.example.csv` to `gp_map.csv`, either in the input folder or in `app/`, or pass
`--gp-map path.csv`. Each row is `pattern,GP name`; the pattern is matched against the fund ID,
fund name and file name, with `*` and `?` as wildcards, and the first match wins. Without a
map the page groups by fund name.

## How the score is built

Four weighted dimensions, each starting at 100 and losing points per failed
check:

| Dimension | Weight | What it measures |
| --- | --- | --- |
| Structure | 15 | Whether the table can be found and mapped to the V7 template |
| Completeness | 30 | Whether required values are present for every position |
| Validity and format | 30 | Whether values have the right type, format and code |
| Consistency | 25 | Whether values agree with each other and with the fund totals |

```
S_d   = 100 - sum over checks c in d of ( w_c * f_c )
Score = sum over d of ( W_d * S_d / 100 )
```

`f_c` is the share of applicable rows or cells that failed, so a large file
does not score lower than a small one of the same quality. File-level checks
give 0 or 1. Aggregate checks, such as weights summing to 100%, scale the
deviation beyond tolerance up to a hard limit. A critical check that fails at
all costs at least 0.25, however few rows are involved.

### Gates

| Gate | Trigger | Effect |
| --- | --- | --- |
| G1 | File cannot be opened | Score 0 |
| G2 | No header found, or fewer than 10 template columns recognised | Score 0 |
| G3 | No position rows, or the CIC column not recognised | Score 0 |
| G4 | Template is not V7 | Cap 40 |
| G5 | More than 20% of mandatory columns missing | Cap 49 |
| G6 | A critical check failed | Cap 94, so Accept is out of reach |
| G7 | More than one TPT table in the workbook (another sheet, or a second header block after a gap on the same sheet) | Cap 40, so Reject |

### Ratings

95 and above Accept, 80 to 94 Minor fixes, 50 to 79 Resubmit, below 50 Reject.

## The 24 checks

| ID | Check | Weight | Critical |
| --- | --- | --- | --- |
| S1 | Header sits in row 1, column A | 15 | |
| S2 | Template version is V7 | 15 | yes |
| S3 | All mandatory columns present | 30 | yes |
| S4 | Column names match the template | 15 | |
| S5 | No duplicate or ambiguous columns | 10 | |
| S6 | No unrecognised extra columns | 5 | |
| S7 | Table body holds only positions | 10 | |
| C1 | Mandatory fields populated | 45 | |
| C2 | Conditional fields populated when their trigger applies | 25 | |
| C3 | Instrument identifier and CIC on every position | 15 | yes |
| C4 | Fund-level fields on every row | 10 | |
| C5 | No placeholders used as values | 5 | |
| V1 | Correct data type: numbers not text, no error values | 20 | |
| V2 | Dates valid and in one format | 20 | |
| V3 | Closed-list values valid: CIC, currency, country, flags | 20 | |
| V4 | Identifier format and check digit: ISIN, LEI | 15 | |
| V5 | Numeric range and scale | 15 | |
| V6 | Text hygiene: stray spaces, lost leading zeros, case | 10 | |
| K1 | Fund-level fields identical on every row | 20 | |
| K2 | Valuation weights sum to 100% | 25 | |
| K3 | Market values add up to the fund NAV | 20 | |
| K4 | Quotation and portfolio values agree through one FX rate | 15 | |
| K5 | No duplicated positions | 10 | yes |
| K6 | Valuation date not after the reporting date, and not stale | 10 | |

Weights are per dimension and add up to 100 within each.

## Structural recovery

A file whose table does not start in A1 is still read. The loader finds the
header row, skips title and instruction rows, drops blank rows, totals rows,
notes and repeated headers, and maps columns by field number, then by template
name, then by a table of GP short labels, then by fuzzy match. Everything it
repairs is written to the recovery log on the file's scorecard and charged
under Structure, so the content is still checked.

Files with no readable header, or a transposed table, fall to gate G2.

## Rules and configuration

The mandatory, conditional and optional flags and the CIC applicability come
straight from the V7 workbook: column K and columns L to AA. Nothing is
hand-maintained. A condition the parser cannot read, such as free text, is
reported as "not evaluated" and carries no penalty.

Every weight, tolerance and threshold lives in `rules.yaml`. Pass it with
`--config rules.yaml`. Keys left out fall back to the built-in default, and
the rules version is printed in every report so a score can be traced to the
settings that produced it.

## Tests

```
python tests/make_test_data.py      # build the fixtures (optional; the tests do it)
python tests/test_tpt_score.py      # 87 tests
```

`tests/data/` holds 33 fixtures: one clean file that scores exactly 100.0, one
file per check carrying a single fault, and one per gate. `expected.json`
records what each should do, so a failing test names the check that broke.

The suite also covers the value parsers against real ISINs and LEIs, the
scoring arithmetic against the worked example in the proposal (93.852), the
rating bands, monotonicity, repeatability, and an end-to-end run that checks
the three reports are produced.

`TestRegressions` pins the defects found in review, each named after the
behaviour that was wrong: trailing blank rows charged as non-position rows, a
zero NAV skipping its check, a cover page named "Data" hiding the positions,
inf and nan accepted as numbers, empty Excel sheet names, and the duplicate
check running quadratically.

The last test case is a calibration run over the sample folder (`../data`) if
it is present: clean baselines must be accepted, wrong-version files
rejected, and unusable files must score 0. Point it elsewhere with
`TPT_SAMPLES=/path python tests/test_tpt_score.py`.

## Layout

```
<project>/
  app/                  everything below
  data/                 the TPT files to score
  output/               where each run writes its reports
  templates/            the FinDatEx V7 workbook

app/
tpt_score.py            entry point
rules.yaml              weights, tolerances, thresholds
tpt_score/
  cli.py                argument parsing and the run
  config.py             configuration and its defaults
  rules.py              field definitions read from the V7 workbook
  loader.py             finds the table, drops non-position rows
  mapper.py             columns to field numbers, version detection
  util.py               value parsing, check digits, reference lists
  scorer.py             weights, floors, gates, bands
  remarks.py            plain-language remarks and the gate reference table
  reporters/management.py  the management page and run_stats.json
  profile.py            per-column data-type mismatches and date formats
  checks/               the 24 checks, one module per dimension
  reporters/            Excel workbook, HTML summary, CSV, console recap
tests/
  make_test_data.py     builds the fixtures
  test_tpt_score.py     the test suite
  data/                 the fixtures and expected.json
```

Nothing outside `output/` is ever written, and no submission file is modified.

Adding a check means writing one function with the `@check('X9', 'title')`
decorator in the right `checks/` module and giving it a weight in
`rules.yaml`. The registry, reports and tests pick it up from there.
