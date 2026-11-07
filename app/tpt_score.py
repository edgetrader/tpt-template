#!/usr/bin/env python3
"""Entry point: score a folder of TPT submissions.

    python tpt_score.py --in <folder with the GP files> [--out <folder>]

Run with --help for every option.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from tpt_score.cli import run        # noqa: E402

if __name__ == '__main__':
    sys.exit(run())
