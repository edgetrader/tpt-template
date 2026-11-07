"""Value parsing and reference data shared by the checks."""
from __future__ import annotations

import datetime as dt
import math
import re
from typing import Optional, Tuple

EXCEL_ERRORS = {'#DIV/0!', '#REF!', '#VALUE!', '#NAME?', '#N/A', '#NULL!', '#NUM!', '#SPILL!'}

PLACEHOLDERS = {'n/a', 'na', 'n.a.', 'n.a', '-', '--', '---', 'null', 'none', 'tbd', 'tba',
                '?', '??', 'xxx', 'nil', 'not available', 'not applicable', '.'}

_THOUSANDS = re.compile(r'^-?[\d.,\s \']+$')
_CURRENCY_PREFIX = re.compile(r'^\s*[-+]?\s*(?:[\$€£¥]|USD|EUR|GBP|CHF|JPY|SGD)\s*', re.I)
_SUFFIX_MULT = {'k': 1e3, 'm': 1e6, 'bn': 1e9, 'b': 1e9}

ISO_CURRENCIES = {
    'AED', 'ARS', 'AUD', 'BGN', 'BHD', 'BRL', 'CAD', 'CHF', 'CLP', 'CNH', 'CNY', 'COP', 'CZK',
    'DKK', 'EGP', 'EUR', 'GBP', 'HKD', 'HRK', 'HUF', 'IDR', 'ILS', 'INR', 'ISK', 'JPY', 'KRW',
    'KWD', 'MAD', 'MXN', 'MYR', 'NGN', 'NOK', 'NZD', 'OMR', 'PEN', 'PHP', 'PKR', 'PLN', 'QAR',
    'RON', 'RSD', 'RUB', 'SAR', 'SEK', 'SGD', 'THB', 'TRY', 'TWD', 'UAH', 'USD', 'VND', 'ZAR',
}

ISO_COUNTRIES = {
    'AE', 'AR', 'AT', 'AU', 'BE', 'BG', 'BH', 'BM', 'BR', 'CA', 'CH', 'CL', 'CN', 'CO', 'CY',
    'CZ', 'DE', 'DK', 'EE', 'EG', 'ES', 'FI', 'FR', 'GB', 'GG', 'GR', 'HK', 'HR', 'HU', 'ID',
    'IE', 'IL', 'IM', 'IN', 'IS', 'IT', 'JE', 'JP', 'KR', 'KW', 'KY', 'LI', 'LT', 'LU', 'LV',
    'MA', 'MT', 'MU', 'MX', 'MY', 'NG', 'NL', 'NO', 'NZ', 'OM', 'PA', 'PE', 'PH', 'PK', 'PL',
    'PT', 'QA', 'RO', 'RS', 'RU', 'SA', 'SE', 'SG', 'SI', 'SK', 'TH', 'TR', 'TW', 'UA', 'US',
    'VG', 'VN', 'ZA',
    'XL', 'XT', 'XX',           # CIC pseudo-countries for non-listed / other
}

CIC_CATEGORIES = set('0123456789ABCDEF')


def text_of(v) -> str:
    return '' if v is None else str(v).strip()


def is_blank(v) -> bool:
    return text_of(v) == ''


def is_excel_error(v) -> bool:
    return text_of(v).upper() in EXCEL_ERRORS


def is_placeholder(v, extra=None) -> bool:
    s = text_of(v).lower()
    if not s:
        return False
    pool = set(PLACEHOLDERS)
    if extra:
        pool |= {str(x).lower() for x in extra}
    return s in pool


def as_number(v) -> Tuple[Optional[float], bool]:
    """Return (value, stored_as_text).

    Recognises thousands separators, decimal commas, currency symbols,
    parentheses for negatives, trailing % and K/M/BN multipliers.
    """
    if v is None:
        return None, False
    if isinstance(v, bool):
        return None, False
    if isinstance(v, (int, float)):
        # inf and nan are not usable numbers: they would poison every sum they
        # reach, so they are reported as a type failure instead.
        return (float(v), False) if math.isfinite(v) else (None, False)
    s = text_of(v)
    if not s or is_excel_error(s) or is_placeholder(s):
        return None, False

    neg = False
    if s.startswith('(') and s.endswith(')'):
        neg, s = True, s[1:-1].strip()
    s = _CURRENCY_PREFIX.sub('', s)
    s = s.replace(' ', ' ').replace("'", '').strip()

    mult = 1.0
    m = re.search(r'([kKmM]|[bB][nN]?)\s*$', s)
    if m and re.match(r'^-?[\d.,\s]+[kKmMbB][nN]?$', s):
        mult = _SUFFIX_MULT.get(m.group(1).lower(), 1.0)
        s = s[:m.start()].strip()
    if s.endswith('%'):
        s = s[:-1].strip()

    if not s or not _THOUSANDS.match(s):
        return None, False
    body = s.replace(' ', '')
    if ',' in body and '.' in body:
        # The rightmost separator is the decimal one.
        if body.rfind(',') > body.rfind('.'):
            body = body.replace('.', '').replace(',', '.')
        else:
            body = body.replace(',', '')
    elif ',' in body:
        frac = body.split(',')[-1]
        body = body.replace(',', '.') if len(frac) != 3 else body.replace(',', '')
    try:
        out = (-1 if neg else 1) * float(body) * mult
    except ValueError:
        return None, False
    return (out, True) if math.isfinite(out) else (None, False)


_DATE_PATTERNS = [
    ('ISO', re.compile(r'^(\d{4})-(\d{2})-(\d{2})$'), (1, 2, 3)),
    ('ISO-slash', re.compile(r'^(\d{4})/(\d{2})/(\d{2})$'), (1, 2, 3)),
    ('compact', re.compile(r'^(\d{4})(\d{2})(\d{2})$'), (1, 2, 3)),
    ('DMY-slash', re.compile(r'^(\d{1,2})/(\d{1,2})/(\d{4})$'), (3, 2, 1)),
    ('DMY-dot', re.compile(r'^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$'), (3, 2, 1)),
    ('DMY-dash', re.compile(r'^(\d{1,2})-(\d{1,2})-(\d{4})$'), (3, 2, 1)),
]


def as_date(v) -> Tuple[Optional[dt.date], str]:
    """Return (date, format name). Format 'none' means it is not a date."""
    if v is None:
        return None, 'none'
    if isinstance(v, dt.datetime):
        return v.date(), 'excel'
    if isinstance(v, dt.date):
        return v, 'excel'
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        n = int(v)
        if 20000 <= n <= 60000:                       # Excel serial day number
            try:
                return (dt.date(1899, 12, 30) + dt.timedelta(days=n)), 'serial'
            except (OverflowError, ValueError):
                return None, 'none'
        if 19000101 <= n <= 21001231:
            v = str(n)
        else:
            return None, 'none'

    s = text_of(v)
    if not s:
        return None, 'none'
    s = s.split(' ')[0].split('T')[0]
    for name, rx, (yi, mi, di) in _DATE_PATTERNS:
        m = rx.match(s)
        if not m:
            continue
        try:
            y, mo, d = int(m.group(yi)), int(m.group(mi)), int(m.group(di))
            if y < 100:
                y += 2000
            # An ambiguous DD/MM vs MM/DD is reported under its own name so the
            # mixed-format check can still see that the file is inconsistent.
            if name.startswith('DMY') and mo > 12 >= d:
                y, mo, d = y, d, mo
                name = name.replace('DMY', 'MDY')
            return dt.date(y, mo, d), name
        except ValueError:
            return None, 'none'
    return None, 'none'


def valid_isin(code) -> bool:
    s = text_of(code).upper().replace(' ', '')
    if not re.fullmatch(r'[A-Z]{2}[A-Z0-9]{9}\d', s):
        return False
    digits = ''.join(str(ord(c) - 55) if c.isalpha() else c for c in s)
    # Luhn: walking from the right, the check digit itself is not doubled.
    total, double = 0, False
    for ch in reversed(digits):
        d = int(ch)
        if double:
            d *= 2
            if d > 9:
                d -= 9
        total += d
        double = not double
    return total % 10 == 0


def valid_lei(code) -> bool:
    s = text_of(code).upper().replace(' ', '')
    if not re.fullmatch(r'[A-Z0-9]{18}\d{2}', s):
        return False
    digits = ''.join(str(ord(c) - 55) if c.isalpha() else c for c in s)
    return int(digits) % 97 == 1


def valid_cic(code) -> bool:
    s = text_of(code).upper().replace(' ', '')
    if len(s) != 4:
        return False
    return s[:2] in ISO_COUNTRIES and s[2] in CIC_CATEGORIES and s[3].isalnum()


def cic_category(code) -> Optional[str]:
    s = text_of(code).upper().replace(' ', '')
    if len(s) == 4 and s[2] in CIC_CATEGORIES:
        return s[2]
    return None


def has_text_noise(v) -> bool:
    """Stray leading/trailing spaces, doubled inner spaces, or a non-breaking space."""
    if not isinstance(v, str) or not v.strip():
        return False
    return v != v.strip() or '  ' in v or ' ' in v or '\n' in v or '\t' in v
