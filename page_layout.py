"""Extract live column widths, row heights and merged ranges from the online
page structure, converting pixels to Excel units and mapping visible columns
to absolute ones through the omitted-columns list. The saved probe file
(edge_layout_probe.json) contains structure only, never cell content."""
import json
from pathlib import Path

SCRIPT_PATH = Path(__file__).resolve().parent / 'web' / 'layout_extract.js'


def script_source():
    return SCRIPT_PATH.read_text(encoding='utf-8')


def px_to_excel_width(px):
    if px is None or px <= 0:
        return None
    return max(0.1, min(255.0, round(px / 7.0, 2)))


def px_to_excel_height(px):
    if px is None or px <= 0:
        return None
    return max(1.0, min(409.0, round(px * 0.75, 2)))


def visible_to_absolute(first_col, omitted, visible_count):
    """Map visible column slots to absolute column numbers, skipping the
    explicitly omitted (hidden) columns."""
    mapping = []
    number = first_col
    omitted_set = set(omitted)
    while len(mapping) < visible_count:
        if number not in omitted_set:
            mapping.append(number)
        number += 1
    return mapping


def extract(frame, root, config, origin, width, total_rows):
    """Read layout from the live grid frame. width/total_rows may be None to
    accept whatever the page rendered."""
    from omitted_columns import omitted_in_range
    from update_excel import address
    first_col, first_row = address(origin.split(':')[0])
    omitted = omitted_in_range(config, first_col, width or 1)
    raw = frame.evaluate(script_source() + '; __extractSheetLayout()')
    (Path(root) / 'edge_layout_probe.json').write_text(json.dumps(raw, ensure_ascii=False), encoding='utf-8')
    if not isinstance(raw, dict) or not raw.get('columns') or not raw.get('rows'):
        raise ValueError('页面结构中找不到可识别的表格网格（列宽/行高），结构已保存到 edge_layout_probe.json，请反馈该文件以便适配')
    mapping = visible_to_absolute(first_col, omitted, len(raw['columns']))
    column_widths = {}
    for index, item in enumerate(raw['columns']):
        value = px_to_excel_width(item.get('width'))
        if value is None:
            continue
        absolute = mapping[index]
        if width is None or first_col <= absolute < first_col + width:
            column_widths[absolute] = value
    row_heights = {}
    base = int(raw.get('firstRow') or 1)
    for index, item in enumerate(raw['rows']):
        value = px_to_excel_height(item.get('height'))
        if value is None:
            continue
        absolute = base + index
        if total_rows is None or first_row <= absolute < first_row + total_rows:
            row_heights[absolute] = value
    end_col = first_col + (width or len(mapping)) - 1
    end_row = first_row + (total_rows or len(raw['rows'])) - 1
    merges = []
    for item in raw.get('merges') or []:
        try:
            left = mapping[int(item['left']) - 1]
            right = mapping[int(item['right']) - 1]
            top = first_row + int(item['top']) - 1
            bottom = first_row + int(item['bottom']) - 1
        except (KeyError, IndexError, TypeError, ValueError):
            continue
        if width is not None:
            right = min(right, end_col)
        if total_rows is not None:
            bottom = min(bottom, end_row)
        if right < left or bottom < top or (left == right and top == bottom):
            continue
        merges.append((left, top, right, bottom))
    probe = raw.get('probe') or {}
    partial = bool(probe.get('virtualized')) or raw.get('strategy') in ('grid', 'headers')
    return dict(sheet='在线页面', column_widths=column_widths, row_heights=row_heights,
                merges=merges, pane=None, partial=partial, strategy=raw.get('strategy'))
