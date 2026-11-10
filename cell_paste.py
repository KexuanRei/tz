"""Plan rectangles containing only changed cells, never unchanged gaps."""
from upload_verify import values_equal


def changed_spans(wanted, actual, first_col, skipped=(), forced=()):
    selected = [first_col + i for i, value in enumerate(wanted)
                if first_col + i not in skipped and
                (first_col + i in forced or not values_equal(value, actual[i]))]
    spans = []
    for col in selected:
        if spans and spans[-1][1] + 1 == col:
            spans[-1] = (spans[-1][0], col)
        else:
            spans.append((col, col))
    return spans


def plan_groups(entries, actual_rows, first_col, skipped=(), forced=None):
    """Merge adjacent target rows only if they have the very same changed span."""
    spans = {}
    for entry in entries:
        for left, right in changed_spans(entry['row'], actual_rows[entry['pk']], first_col,
                                         skipped, (forced or {}).get(entry['pk'], ())):
            spans.setdefault((left, right), []).append(entry)
    groups = []
    for (left, right), items in sorted(spans.items()):
        group = []
        for entry in items:
            if group and (entry['target'] != group[-1]['target'] + 1 or
                          entry['source'] != group[-1]['source'] + 1):
                groups.append((group, left, right))
                group = []
            group.append(entry)
        if group:
            groups.append((group, left, right))
    return groups


def verify_written_formats(copier, page, ref, expected, formats, keys, max_cells):
    """Retry/split real read-backs without ever repeating a paste."""
    from clipboard_html import CopyIntegrityError
    from html_cell_formats import FormatReadbackError
    from sync_runner import _stop_requested, SyncStopped
    from upload_verify import snapshot, differences
    from update_excel import address
    from copy_range import column_name
    left, top = address(ref.split(':')[0])
    right, bottom = address(ref.split(':')[1])
    queue = [(left, top, right, bottom, 0)]
    while queue:
        if _stop_requested(copier):
            raise SyncStopped()
        a, y, b, z, attempt = queue.pop()
        area = f'{column_name(a)}{y}:{column_name(b)}{z}'
        try:
            clip = copier._capture(page, area)
            offset = None
            if clip.get('blank'):
                reader = getattr(formats, 'capture_blank_formats', None)
                if reader is None:
                    raise FormatReadbackError('空白复制不含真实格式，保留待重试状态')
                clip, actual, offset = reader(page, area, max_cells)
            else:
                actual = snapshot(clip, area, max_cells, strict_rows=True)
            wanted = dict(rows=[row[a-left:b-left+1] for row in expected['rows'][y-top:z-top+1]],
                          row_count=z-y+1, width=b-a+1)
            if differences(wanted, actual, area.split(':')[0], limit=1):
                raise ValueError('表A在格式校验前内容发生变化，保留待同步状态')
            options = {} if offset is None else {'row_offset': offset}
            formats.verify(clip, keys[y-top:z-top+1], a, b, **options)
        except CopyIntegrityError as exc:
            if exc.can_split and exc.kind == 'columns' and a < b:
                mid = (a+b)//2
                queue.extend([(mid+1, y, b, z, 0), (a, y, mid, z, 0)])
                continue
            if exc.can_split and exc.kind != 'columns' and y < z:
                mid = (y+z)//2
                queue.extend([(a, mid+1, b, z, 0), (a, y, b, mid, 0)])
                continue
            raise
        except FormatReadbackError:
            if attempt >= 3:
                raise
            copier.update(message=f'正在等待表A {area} 格式回读稳定（仅重新读取，不重复粘贴）…')
            page.wait_for_timeout(500 * (attempt+1))
            queue.append((a, y, b, z, attempt+1))
