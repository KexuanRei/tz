"""Plan table C's layout and merge-safe uploads without changing a workbook.

Excel column widths are character-based, not pixels.  ``value`` preserves that
unit; ``points`` is optional geometry measured by Excel COM from table C.  A web
adapter may convert those measured points at its CSS DPI, but must not assume
that dividing or multiplying an Excel width by seven reproduces the layout.
"""
import math

from copy_range import column_name
from update_excel import address


MAX_COLUMNS = 16384
MAX_ROWS = 1048576
FALLBACK_ROW_HEIGHT = 15.0
FALLBACK_COLUMN_WIDTH = 8.43


def _integer(value, label, maximum):
    if type(value) is not int or not 1 <= value <= maximum:
        raise ValueError(f'{label}必须是1至{maximum}的整数')
    return value


def _number(value, label, *, allow_zero=False):
    if isinstance(value, bool):
        raise ValueError(f'{label}必须是有效尺寸')
    try:
        result = float(value)
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError(f'{label}必须是有效尺寸') from exc
    if not math.isfinite(result) or result < 0 or (result == 0 and not allow_zero):
        raise ValueError(f'{label}必须是有效尺寸')
    return result


def _bounds(ref):
    if not isinstance(ref, str):
        raise ValueError('上传源区域必须是A1:B20这样的单元格范围')
    parts = ref.strip().upper().split(':')
    if len(parts) not in (1, 2):
        raise ValueError('上传源区域必须是A1:B20这样的单元格范围')
    left, top = address(parts[0])
    right, bottom = address(parts[-1])
    if right < left or bottom < top:
        raise ValueError('上传源区域的结束位置不能在起始位置之前')
    return left, top, right, bottom


def _ref(bounds):
    left, top, right, bottom = bounds
    return f'{column_name(left)}{top}:{column_name(right)}{bottom}'


def _contained_merges(merges, bounds):
    left, top, right, bottom = bounds
    result = []
    seen = set()
    for item in merges or ():
        if isinstance(item, str):
            merge = _bounds(item)
        else:
            try:
                merge = tuple(item)
            except TypeError as exc:
                raise ValueError('表C包含无效的合并区域') from exc
            if len(merge) != 4:
                raise ValueError('表C包含无效的合并区域')
            for coordinate, maximum in zip(merge, (MAX_COLUMNS, MAX_ROWS, MAX_COLUMNS, MAX_ROWS)):
                _integer(coordinate, '合并区域坐标', maximum)
        c1, r1, c2, r2 = merge
        if c1 > c2 or r1 > r2:
            raise ValueError('表C包含起止位置颠倒的合并区域')
        if c2 < left or c1 > right or r2 < top or r1 > bottom:
            continue
        if c1 < left or c2 > right or r1 < top or r2 > bottom:
            raise ValueError(f'上传源区域{_ref(bounds)}未完整包含表C的合并区域{_ref(merge)}，请扩大上传范围')
        if (c1, r1) != (c2, r2) and merge not in seen:
            result.append(merge)
            seen.add(merge)
    return sorted(result, key=lambda item: (item[1], item[0], item[3], item[2]))


def _destination(bounds, destination):
    col, row = address(destination.strip().upper() if isinstance(destination, str) else destination)
    left, top, right, bottom = bounds
    end_col, end_row = col + right - left, row + bottom - top
    if end_col > MAX_COLUMNS or end_row > MAX_ROWS:
        raise ValueError('表C粘贴区域超出工作表行列范围')
    return col, row, end_col, end_row


def translate_merges(merges, source_range, destination='A1'):
    """Map complete source merges to table C, rejecting clipped merges."""
    bounds = _bounds(source_range)
    target = _destination(bounds, destination)
    dc, dr = target[0] - bounds[0], target[1] - bounds[1]
    return [_ref((c1 + dc, r1 + dr, c2 + dc, r2 + dr))
            for c1, r1, c2, r2 in _contained_merges(merges, bounds)]


def _index_map(values, label):
    if values is None:
        return {}
    if not isinstance(values, dict):
        raise ValueError(f'{label}必须按原工作表行列号提供')
    result = {}
    for key, value in values.items():
        if type(key) is int:
            index = key
        elif isinstance(key, str) and key.isascii() and key.isdigit():
            index = int(key)
        else:
            raise ValueError(f'{label}包含无效行列号')
        if index < 1:
            raise ValueError(f'{label}包含无效行列号')
        result[index] = _number(value, label, allow_zero=True)
    return result


def group_dimensions(values, first, last, target_first, *, default, unit,
                     default_source='reference_default', points=None, default_points=None):
    """Group adjacent equal dimensions, including all unspecified positions.

    Coordinates in the returned groups belong to table C.  Sparse table C
    dimensions are processed by intervals, so a million default-height rows
    become one operation.  The optional points map is used for columns only.
    """
    limit = MAX_COLUMNS if unit == 'excel_char' else MAX_ROWS
    for value, label in ((first, '源起始位置'), (last, '源结束位置'), (target_first, '目标起始位置')):
        _integer(value, label, limit)
    if last < first or target_first + last - first > limit:
        raise ValueError('行列尺寸应用范围无效或超出工作表范围')
    if unit not in ('excel_char', 'pt'):
        raise ValueError('行列尺寸单位无效')
    default = _number(default, '默认尺寸')
    sizes = _index_map(values, '表C尺寸')
    measures = _index_map(points, '表C实测列宽') if points is not None else {}
    if default_points is not None:
        default_points = _number(default_points, '表C默认列宽实测值', allow_zero=True)
    boundaries = sorted({number for number in sizes.keys() | measures.keys() if first <= number <= last})
    groups = []
    offset = target_first - first

    def append(start, end, value, source, measured):
        if start > end:
            return
        start, end = start + offset, end + offset
        if groups and groups[-1]['end'] + 1 == start and groups[-1]['value'] == value and groups[-1].get('points') == measured:
            group = groups[-1]
            group['end'] = end
            if source not in group['sources']:
                group['sources'].append(source)
                group['source'] = 'mixed'
            return
        group = dict(start=start, end=end, value=value, unit=unit, source=source, sources=[source])
        if unit == 'excel_char':
            group['points'] = measured
        groups.append(group)

    cursor = first
    for number in boundaries:
        append(cursor, number - 1, default, default_source, default_points)
        explicit = number in sizes
        # A default-column measurement cannot stand in for a custom width.
        measured = measures.get(number, None if explicit else default_points)
        append(number, number, sizes.get(number, default), 'explicit' if explicit else default_source, measured)
        cursor = number + 1
    append(cursor, last, default, default_source, default_points)
    return groups


def _default(reference, key, caller, fallback):
    if reference.get(key) is not None:
        return _number(reference[key], '表C默认尺寸'), 'reference_default'
    if caller is not None:
        return _number(caller, 'Excel默认尺寸'), 'caller_default'
    return fallback, 'fallback'


def _pane(pane, bounds, target):
    if pane is None:
        return None
    if not isinstance(pane, (tuple, list)) or len(pane) != 3:
        raise ValueError('表C冻结窗格信息无效')
    x_split, y_split, top_left_cell = pane
    if any(type(value) is not int or value < 0 for value in (x_split, y_split)):
        raise ValueError('表C冻结窗格位置无效')
    left, top, right, bottom = bounds
    # Freeze lines are worksheet-wide.  Offsetting a frozen source therefore
    # also freezes any rows/columns preceding the pasted region in table C.
    frozen_columns = max(0, min(x_split, right) - left + 1)
    frozen_rows = max(0, min(y_split, bottom) - top + 1)
    if not frozen_columns and not frozen_rows:
        return None
    tx = target[0] + frozen_columns - 1 if frozen_columns else 0
    ty = target[1] + frozen_rows - 1 if frozen_rows else 0
    scroll_col, scroll_row = address(top_left_cell)
    scroll_col = max(tx + 1, target[0] + max(0, scroll_col - left))
    scroll_row = max(ty + 1, target[1] + max(0, scroll_row - top))
    scroll_col = min(MAX_COLUMNS, scroll_col)
    scroll_row = min(MAX_ROWS, scroll_row)
    return dict(x_split=tx, y_split=ty, top_left_cell=f'{column_name(scroll_col)}{scroll_row}')


def plan_layout(reference, source_range, destination='A1', *,
                default_row_height=None, default_col_width=None):
    """Map table C dimensions, merges and freeze positions to table C.

    Caller defaults should be read from Excel's StandardHeight/StandardWidth
    on table C.  15pt and 8.43 characters are explicitly reported fallback
    defaults only; they are not claimed to be measurements of this workbook.
    """
    if not isinstance(reference, dict):
        raise ValueError('请先读取表C的格式信息')
    bounds = _bounds(source_range)
    target = _destination(bounds, destination)
    row_default, row_source = _default(reference, 'default_row_height', default_row_height, FALLBACK_ROW_HEIGHT)
    col_default, col_source = _default(reference, 'default_col_width', default_col_width, FALLBACK_COLUMN_WIDTH)
    columns = group_dimensions(reference.get('column_widths'), bounds[0], bounds[2], target[0],
                               default=col_default, unit='excel_char', default_source=col_source,
                               points=reference.get('column_width_points'),
                               default_points=reference.get('default_col_width_points'))
    rows = group_dimensions(reference.get('row_heights'), bounds[1], bounds[3], target[1],
                            default=row_default, unit='pt', default_source=row_source)
    warnings = []
    if any('fallback' in group['sources'] for group in rows):
        warnings.append('表C未提供默认行高，未声明行采用默认值15磅；建议通过Excel读取StandardHeight')
    if any('fallback' in group['sources'] for group in columns):
        warnings.append('表C未提供默认列宽，未声明列采用默认值8.43字符；建议通过Excel读取StandardWidth')
    return dict(source_range=_ref(bounds), target_range=_ref(target), column_groups=columns,
                row_groups=rows, merges=translate_merges(reference.get('merges'), source_range, destination),
                pane=_pane(reference.get('pane'), bounds, target), warnings=warnings)


def plan_batches(source_start, rows, columns, batch_rows, *, merges=(), max_cells=100_000_000):
    """Return ``(source_ref, first_row, count)`` batches without cutting merges.

    Crossing merges extend the current batch, including chains of merge row
    spans.  The maximum cell limit remains hard; a merge that cannot fit fails
    during planning before any online paste is attempted.
    """
    if not isinstance(source_start, (tuple, list)) or len(source_start) != 2:
        raise ValueError('上传起始位置无效')
    first_col = _integer(source_start[0], '上传起始列', MAX_COLUMNS)
    first_row = _integer(source_start[1], '上传起始行', MAX_ROWS)
    _integer(rows, '上传行数', MAX_ROWS)
    _integer(columns, '上传列数', MAX_COLUMNS)
    _integer(batch_rows, '每批行数', MAX_ROWS)
    if type(max_cells) is not int or max_cells < 1:
        raise ValueError('每批最大单元格数必须是正整数')
    end_col, end_row = first_col + columns - 1, first_row + rows - 1
    if end_col > MAX_COLUMNS or end_row > MAX_ROWS:
        raise ValueError('上传源区域超出工作表范围')
    if columns > max_cells:
        raise ValueError('一行数据已超过每批最大单元格数，请提高上限')
    scoped_merges = _contained_merges(merges, (first_col, first_row, end_col, end_row))
    row_spans = sorted({(top, bottom) for _, top, _, bottom in scoped_merges if top < bottom})
    # Merge overlapping vertical spans even when the cells occupy different
    # columns: a whole-row upload cannot split either member of such a chain.
    bands = []
    for top, bottom in row_spans:
        if bands and top <= bands[-1][1]:
            bands[-1] = (bands[-1][0], max(bands[-1][1], bottom))
        else:
            bands.append((top, bottom))
    batches = []
    row = first_row
    band_index = 0
    nominal_rows = min(batch_rows, max_cells // columns)
    while row <= end_row:
        end = min(end_row, row + nominal_rows - 1)
        while band_index < len(bands) and bands[band_index][1] <= end:
            band_index += 1
        if band_index < len(bands) and bands[band_index][0] <= end < bands[band_index][1]:
            band_top, band_bottom = bands[band_index]
            if (band_bottom - row + 1) * columns > max_cells and band_top > row:
                # Leave the complete merge for the next batch rather than
                # failing because unrelated preceding rows used the budget.
                end = band_top - 1
            else:
                end = band_bottom
                band_index += 1
        count = end - row + 1
        if count * columns > max_cells:
            raise ValueError(f'保持合并区域完整需要一次上传{count}行×{columns}列，超过每批最大单元格数{max_cells}；请提高上限或调整源范围')
        batches.append((_ref((first_col, row, end_col, end)), row, count))
        row = end + 1
    return batches
