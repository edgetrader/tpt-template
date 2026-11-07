"""Run reports: Excel workbook, one-page HTML, CSV and console recap."""
from .console import recap, write_csv
from .excel import feedback_text, write_workbook
from .html import write_html

__all__ = ['recap', 'write_csv', 'write_workbook', 'write_html', 'feedback_text']
