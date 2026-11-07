"""Scoring configuration: every weight, tolerance and threshold lives here."""
from __future__ import annotations

import copy
import os
from dataclasses import dataclass, field as dc_field
from typing import Any, Dict

try:
    import yaml
except ImportError:                       # pragma: no cover - yaml ships with the tool
    yaml = None

DEFAULTS: Dict[str, Any] = {
    'rules_version': '2026.1',
    'template': 'templates/TPT_V7  20241125_updated.xlsx',
    'dimensions': {'structure': 15, 'completeness': 30, 'validity': 30, 'consistency': 25},
    'critical_floor': 0.25,
    'bands': {'accept': 95, 'minor_fixes': 80, 'resubmit': 50},
    'gates': {
        'min_mapped_columns': 10,
        'wrong_version_cap': 40,
        'multiple_tables_cap': 40,
        'missing_mandatory': {'share': 0.20, 'cap': 49},
        'critical_failure_cap': 94,
    },
    'checks': {
        # Structure (weights add up to 100 inside the dimension)
        'S1': {'weight': 15, 'critical': False},
        'S2': {'weight': 15, 'critical': True},
        'S3': {'weight': 30, 'critical': True},
        'S4': {'weight': 15, 'critical': False},
        'S5': {'weight': 10, 'critical': False},
        'S6': {'weight': 5, 'critical': False},
        'S7': {'weight': 10, 'critical': False},
        # Completeness
        'C1': {'weight': 45, 'critical': False},
        'C2': {'weight': 25, 'critical': False},
        'C3': {'weight': 15, 'critical': True, 'fields': ['12', '14']},
        'C4': {'weight': 10, 'critical': False},
        'C5': {'weight': 5, 'critical': False,
               'placeholders': ['n/a', 'na', 'n.a.', '-', '--', 'null', 'none', 'tbd', 'tba', '?', 'xxx', 'nil']},
        # Validity and format
        'V1': {'weight': 20, 'critical': False},
        'V2': {'weight': 20, 'critical': False},
        'V3': {'weight': 20, 'critical': False},
        'V4': {'weight': 15, 'critical': False},
        'V5': {'weight': 15, 'critical': False},
        'V6': {'weight': 10, 'critical': False},
        # Consistency
        'K1': {'weight': 20, 'critical': False},
        'K2': {'weight': 25, 'critical': False, 'tolerance_pp': 0.5, 'hard_limit_pp': 5.0},
        'K3': {'weight': 20, 'critical': False, 'tolerance_pct': 0.5, 'hard_limit_pct': 5.0},
        'K4': {'weight': 15, 'critical': False, 'tolerance_pct': 1.0},
        'K5': {'weight': 10, 'critical': True},
        'K6': {'weight': 10, 'critical': False, 'stale_days': 100},
    },
    'reports': {'file_scorecard_sheets_max': 150, 'html': True, 'excel': True},
}


@dataclass
class Config:
    data: Dict[str, Any] = dc_field(default_factory=lambda: copy.deepcopy(DEFAULTS))
    source: str = '(built-in defaults)'

    # -- convenience accessors --------------------------------------------
    @property
    def rules_version(self) -> str:
        return str(self.data.get('rules_version', ''))

    @property
    def dimensions(self) -> Dict[str, float]:
        return self.data['dimensions']

    @property
    def bands(self) -> Dict[str, float]:
        return self.data['bands']

    @property
    def gates(self) -> Dict[str, Any]:
        return self.data['gates']

    @property
    def critical_floor(self) -> float:
        return float(self.data.get('critical_floor', 0.25))

    def check(self, check_id: str) -> Dict[str, Any]:
        return self.data['checks'].get(check_id, {})

    def weight(self, check_id: str) -> float:
        return float(self.check(check_id).get('weight', 0))

    def is_critical(self, check_id: str) -> bool:
        return bool(self.check(check_id).get('critical', False))

    def rating(self, score: float) -> str:
        b = self.bands
        if score >= b['accept']:
            return 'Accept'
        if score >= b['minor_fixes']:
            return 'Minor fixes'
        if score >= b['resubmit']:
            return 'Resubmit'
        return 'Reject'


def _deep_merge(base: Dict[str, Any], patch: Dict[str, Any]) -> Dict[str, Any]:
    out = copy.deepcopy(base)
    for k, v in (patch or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path: str = None) -> Config:
    """Load ``rules.yaml`` over the built-in defaults."""
    if not path:
        return Config()
    if not os.path.exists(path):
        raise FileNotFoundError('config file not found: %s' % path)
    if yaml is None:
        raise RuntimeError('PyYAML is required to read %s' % path)
    with open(path, 'r', encoding='utf-8') as fh:
        patch = yaml.safe_load(fh) or {}
    return Config(data=_deep_merge(DEFAULTS, patch), source=path)


def validate(cfg: Config) -> list:
    """Return a list of complaints about the configuration, empty when sound."""
    problems = []
    dims = cfg.dimensions
    total = sum(float(v) for v in dims.values())
    if abs(total - 100.0) > 1e-6:
        problems.append('dimension weights add up to %g, not 100' % total)
    groups = {'structure': 'S', 'completeness': 'C', 'validity': 'V', 'consistency': 'K'}
    for dim, prefix in groups.items():
        s = sum(float(v.get('weight', 0)) for k, v in cfg.data['checks'].items()
                if k.startswith(prefix))
        if abs(s - 100.0) > 1e-6:
            problems.append('%s check weights add up to %g, not 100' % (dim, s))
    b = cfg.bands
    if not (b['resubmit'] < b['minor_fixes'] < b['accept'] <= 100):
        problems.append('rating bands are not in ascending order')
    return problems
