"""Read a chosen header row without assuming merged copies stay rectangular."""
import csv
import io
import re

from lxml import html

from clipboard_html import CopyIntegrityError, table_fragment
from copy_range import column_name
from upload_verify import snapshot


def header_name(value):
    # A wrapped heading is still one name, e.g. 分段< br >上胎.
    return re.sub(r'[\s\u200b\ufeff]+', '', str(value if value is not None else ''))


def single_header(clip, ref, max_cells, copied_ref=None):
    """Accept exactly one logical cell, even when it spans several rows/columns.

    The browser selects ref by its absolute address before copying. We never
    infer an absolute column from a compressed wide clipboard's column count.
    """
    fragment = table_fragment(clip['html'])
    table = html.fromstring(fragment)
    cells = table.xpath('.//td|.//th')
    if len(cells) != 1:
        raise ValueError(f'表C表头 {ref} 复制后包含{len(cells)}个单元格，无法确认标题，请检查该处合并或隐藏状态')
    cell = cells[0]
    try:
        width, height = int(cell.get('colspan', '1')), int(cell.get('rowspan', '1'))
        if width < 1 or height < 1 or width * height > max_cells:
            raise ValueError()
    except (ValueError, TypeError) as exc:
        raise ValueError(f'表C表头 {ref} 合并范围无效或超过读取上限') from exc
    if copied_ref is not None:
        from update_excel import address
        target_col, target_row = address(ref.split(':')[0])
        left, top = address(copied_ref.split(':')[0])
        if not (left <= target_col < left+width and top <= target_row < top+height):
            raise ValueError(f'表C表头 {ref} 不在实际复制的合并单元格内，已停止筛选')
    name = ''.join(cell.itertext())
    # TSV can contain blank rows occupied by a rowspan, or newlines inside a
    # quoted header. A horizontal merge (colspan) repeats the title across its
    # columns (e.g. 预组吨位 → 预组吨位预组吨位). Take the first non-empty value
    # as the title and ignore the merge-expansion placeholders.
    try:
        values = [value for row in csv.reader(io.StringIO(clip['text']), delimiter='\t') for value in row]
    except (KeyError, csv.Error) as exc:
        raise ValueError(f'表C表头 {ref} 缺少有效文本，无法确认标题') from exc
    plain = next((value for value in values if value.strip()), '')
    if header_name(name) != header_name(plain):
        raise ValueError(f'表C表头 {ref} 富文本与文本标题不一致，已停止筛选'
                         f'（富文本={name!r}，纯文本={plain!r}，'
                         f'归一化后=富文本{header_name(name)!r} 对 纯文本{header_name(plain)!r}）')
    return name


def capture_header_cell(copier, page, ref, max_cells):
    from range_selection import RangeSelectionError
    from update_excel import address
    try:
        return single_header(copier._capture(page, ref), ref, max_cells)
    except RangeSelectionError as exc:
        # WPS may normalize A4 to the anchor A2 (or A2:A4). Only this
        # read-only header path can follow that observed address. The copied
        # single cell's span must then prove it actually covers A4.
        detail = exc.diagnostic_context
        requested = str(detail.get('requested_range') or '').replace('$','').upper()
        current = str(detail.get('control',{}).get('address') or '').replace('$','').upper()
        if exc.code != 'RANGE_NOT_CONFIRMED' or requested not in (ref, ref.split(':')[0]) or not current:
            raise
        try:
            left, top = address(current.split(':')[0])
            target_col, target_row = address(ref.split(':')[0])
            if left > target_col or top > target_row:
                raise ValueError()
            if ':' in current:
                right, bottom = address(current.split(':')[1])
                if right < target_col or bottom < target_row:
                    raise ValueError()
        except ValueError:
            raise exc
        copied_ref = current if ':' in current else f'{current}:{current}'
        return single_header(copier._capture(page, copied_ref), ref, max_cells, copied_ref)


def read_header_row(copier, page, first_col, last_col, row, max_cells, skipped=(), stop_names=None):
    from sync_runner import _stop_requested, SyncStopped
    from range_selection import RangeSelectionError
    if _stop_requested(copier):
        raise SyncStopped()
    skipped = set(skipped)
    ref = f'{column_name(first_col)}{row}:{column_name(last_col)}{row}'
    # Ordinary single-level headers retain the fast one-copy path. A skipped column
    # (e.g. a hidden sync exclusion) must stay blank at its original position, so a
    # full-range copy cannot be trusted when any column is skipped.
    if not skipped:
        try:
            result = snapshot(copier._capture(page, ref), ref, max_cells, strict_rows=True)
            if not result['merges']:
                return result['rows'][0]
        except CopyIntegrityError:
            pass
        except RangeSelectionError as exc:
            if exc.code != 'RANGE_NOT_CONFIRMED':
                raise
    names = []
    remaining = set(stop_names) if stop_names else None
    for col in range(first_col, last_col + 1):
        if _stop_requested(copier):
            raise SyncStopped()
        if col in skipped:
            names.append('')
            continue
        address = f'{column_name(col)}{row}'
        copier.update(message=f'正在逐列识别第{row}行合并表头：{address}（{col-first_col+1}/{last_col-first_col+1}）…')
        cell_ref = f'{address}:{address}'
        name = capture_header_cell(copier, page, cell_ref, max_cells)
        names.append(name)
        if remaining is not None:
            remaining.discard(header_name(name))
            if not remaining:
                break
    return names
