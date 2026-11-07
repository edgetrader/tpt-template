"""The 24 checks. Importing this package registers all of them."""
from .base import (CHECKS, CheckResult, CheckSpec, Context, Issue, band, check, run_all,
                   share)
from . import structure, completeness, validity, consistency   # noqa: F401  (registration)

__all__ = ['CHECKS', 'CheckResult', 'CheckSpec', 'Context', 'Issue', 'band', 'check',
           'run_all', 'share']
