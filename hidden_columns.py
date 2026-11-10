"""Auto-detect hidden columns from the online sheet's column header strip.

The page keeps a frozen row of column letters (A, B, C …).  A hidden column has
no letter there, so the letters inside the requested range show exactly which
columns were hidden.  Detection is deliberately conservative: it is only used
when both ends of the requested range were rendered, and the copy itself still
has to match the resulting width before anything is pasted.
"""
from pathlib import Path

from copy_range import column_name
from omitted_columns import column_list, parse_omitted_columns
from update_excel import column

SCRIPT_PATH = Path(__file__).resolve().parent / 'web' / 'hidden_columns.js'


def script_source():
    return SCRIPT_PATH.read_text(encoding='utf-8')


def detect(page, first_col, last_col):
    """Return (hidden_column_numbers, detail). Never raises; () means 'unknown'."""
    if last_col < first_col:
        return (), {'reason': 'bad-range'}
    script = script_source()
    call = '__detectHiddenColumns(%r, %r)' % (column_name(first_col), column_name(last_col))
    from range_selection import _frames
    last_detail = {}
    for frame in _frames(page):
        try:
            raw = frame.evaluate(script + ';' + call)
        except Exception:
            continue
        if not isinstance(raw, dict):
            continue
        last_detail = raw
        if not raw.get('ok'):
            continue
        hidden = []
        for letter in raw.get('missing') or ():
            try:
                hidden.append(column(str(letter)))
            except (TypeError, ValueError):
                return (), {'reason': 'bad-letter', 'detail': raw}
        hidden = tuple(sorted({c for c in hidden if first_col <= c <= last_col}))
        return hidden, raw
    return (), last_detail


def with_omitted(config, columns):
    """Return a copy of config whose omitted columns include the auto-detected ones."""
    columns = tuple(sorted(set(columns)))
    if not columns:
        return config
    merged = set(parse_omitted_columns(config['source'].get('omitted_columns', '')))
    merged.update(columns)
    copy = dict(config)
    copy['source'] = dict(config['source'])
    copy['source']['omitted_columns'] = column_list(sorted(merged))
    return copy


def describe(columns):
    return column_list(columns)
