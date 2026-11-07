#!/usr/bin/env python3
"""Entry point: combine the accepted TPT files of the last run into one table.

    python tpt_concat.py --in <folder that tpt_score.py scored>

Writes tpt_combined.xlsx beside accepted_files.xlsx in output/<folder>/.
Run with --help for every option.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from tpt_score.concat import run        # noqa: E402

if __name__ == '__main__':
    sys.exit(run())
