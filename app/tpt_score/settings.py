"""Local settings: where the submissions and the fund reference live.

``<project>/config/settings.yaml`` holds paths that belong to one machine, so
they need not be typed on every run:

    input_folder: /path/to/submissions        # root folder of the TPT files
    fund_reference: /path/to/fund_reference.xlsx

A relative path is taken from the project folder, ``~`` is the home folder.
Command-line options always win over the file. The file is optional.
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional

import yaml

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DEFAULT_SETTINGS = os.path.join(PROJECT_DIR, 'config', 'settings.yaml')
KEYS = ('input_folder', 'fund_reference')


class SettingsError(Exception):
    pass


@dataclass
class Settings:
    path: str = ''                       # '' when no settings file was read
    input_folder: Optional[str] = None
    fund_reference: Optional[str] = None


def _resolve(value) -> Optional[str]:
    if value is None or not str(value).strip():
        return None
    p = os.path.expanduser(os.path.expandvars(str(value).strip()))
    return p if os.path.isabs(p) else os.path.normpath(os.path.join(PROJECT_DIR, p))


def load_settings(path: Optional[str] = None) -> Settings:
    """Read ``path``, or the default file when it exists. A missing default is fine;
    a missing file that was asked for is an error."""
    target = path or DEFAULT_SETTINGS
    if not os.path.isfile(target):
        if path:
            raise SettingsError('settings file not found: %s' % path)
        return Settings()
    try:
        with open(target, encoding='utf-8') as fh:
            data = yaml.safe_load(fh) or {}
    except yaml.YAMLError as exc:
        raise SettingsError('settings file %s is not valid YAML: %s' % (target, exc))
    if not isinstance(data, dict):
        raise SettingsError('settings file %s should hold key: value lines' % target)
    unknown = sorted(set(data) - set(KEYS))
    if unknown:
        raise SettingsError('settings file %s has unknown keys: %s (known: %s)'
                            % (target, ', '.join(unknown), ', '.join(KEYS)))
    return Settings(path=target, input_folder=_resolve(data.get('input_folder')),
                    fund_reference=_resolve(data.get('fund_reference')))
