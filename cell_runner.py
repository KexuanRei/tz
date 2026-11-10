"""新增记录插入：把表A1 出现、表A2 还没有的主键，按“最近主键锚点 + Δ行”插入表A2。

只处理“新增”：
- 主键已在表A2 中的记录（修改）交给 sync_runner.run_sync 仅复制差异单元格，这里不碰；
- 主键不在表A2 中的新记录，以上下最近的已有主键为锚点，写入表A2 锚点附近的目标行，
  并把新主键一并写入（更新主键）。
- 目标行若原本已有数据，也照常写入（覆盖）：表A1 非空且与表A 不同的单元格会覆盖表A
  原值，含主键列；表A1 为空的单元格仍跳过，不清空表A。这样“表A 原本有数据”时，
  表C 的改动也能落到表A。
"""
from copy_range import column_name
from update_excel import address
from upload_verify import snapshot, differences, normalize_text
from cell_paste import changed_spans, verify_written_formats
from sync_engine import parse_pk_column, data_start_row, build_index, sync_batch_rows
from sync_engine import legacy_scan
from sync_columns import effective_skip_columns, format_skip_columns
from cell_sync import find_inserts, changed_cells

# 复用 sync_runner 的页面打开、范围读取、粘贴校验等基础原语。
from sync_runner import (
    _open_sheet, _groups, _read_kept_columns, _verify_rows, read_online_grid,
    _log_diffs, _new_mirror, _read_row_windows, _stop_requested, SyncStopped,
)


def _build_a_index(copier, config, page_a, pk_number, pk_name, data_start, skipped, c1, c2):
    """把表A2 主键列读成 {主键值: 行号}。与 run_sync 的 target_map 逻辑一致。

    默认（非旧检索方式）不再在线分批读取表A主键列：主键→行号直接来自表B（表A 整表下载快照），
    即“直接扫描全表”。旧检索方式（legacy_scan）仍按 scan_range + batch_rows 在线读取。
    """
    if not legacy_scan(config):
        from sync_runner import read_table_b_source
        from table_b_compare import target_map_from_table_b
        table_b = read_table_b_source(copier, config, c1, c2)
        return target_map_from_table_b(table_b)
    scan_range = config['upload'].get('scan_range') or None
    if not scan_range and skipped:
        scan_range = copier._used_range(page_a)
    if scan_range:
        left, right = scan_range.split(':')
        ac1, ar1 = address(left)
        ac2, ar2 = address(right)
        if not ac1 <= pk_number <= ac2:
            raise ValueError(f'表A扫描区域必须包含主键列 {pk_name}')
        scan_range = f'{pk_name}{ar1}:{pk_name}{ar2}'
    grid = read_online_grid(copier, page_a, config, scan_range, '表A主键')
    first = grid['first_column']
    if not first <= pk_number < first + grid['width']:
        raise ValueError(f'表A扫描区域必须包含主键列 {pk_name}')
    index = build_index(grid, pk_name, pk_number, data_start,
                        config['upload'].get('scan_range') or '', label='表A')
    return {key: row_no for key, (_row, row_no) in index['rows'].items()}


def _read_rows_by_number(copier, page, rows, c1, c2, max_cells, skipped, batch_rows=100):
    """读取指定页（表C/表A）中若干行号的整行，返回 {行号: values}。"""
    result = {}
    for group in _read_row_windows(rows, batch_rows, max_cells, c2 - c1 + 1):
        if _stop_requested(copier):
            raise SyncStopped()
        ref = f'{column_name(c1)}{group[0]}:{column_name(c2)}{group[-1]}'
        block = _read_kept_columns(copier, page, ref, max_cells, skipped)
        for row_no in group:
            result[row_no] = block['rows'][row_no - group[0]]
    return result


def _is_blank_row(values, c1, width, skipped):
    for i in range(width):
        col_abs = c1 + i
        if col_abs in skipped:
            continue
        if normalize_text(values[i] if i < len(values) else ''):
            return False
    return True


def _paste_block(copier, page_c, page_a, source_row, target_row, left_col, right_col,
                 max_cells, expected_values, before_paste=None, formats=None, key=None):
    """把表A1 source_row 的 left_col:right_col 复制并粘贴到表A target_row 同一列区间。"""
    source_ref = f'{column_name(left_col)}{source_row}:{column_name(right_col)}{source_row}'
    target_ref = f'{column_name(left_col)}{target_row}:{column_name(right_col)}{target_row}'
    expected = dict(rows=[expected_values], row_count=1, width=right_col - left_col + 1)
    current = _read_kept_columns(copier, page_a, target_ref, max_cells, frozenset())
    spans = changed_spans(expected_values, current['rows'][0], left_col)
    if not spans:
        return 0
    if current['merges']:
        raise ValueError(f'表A {target_ref} 涉及合并单元格，无法只更新差异单元格')
    if spans != [(left_col, right_col)]:
        return sum(_paste_block(copier, page_c, page_a, source_row, target_row, a, b,
                                max_cells, expected_values[a-left_col:b-left_col+1],
                                before_paste, formats, key) for a, b in spans)
    if formats is None:
        raise ValueError('粘贴差异单元格必须提供表C同主键同列格式，尚未粘贴表A')
    formats.preflight([key], left_col, right_col)
    found = differences(expected, current, target_ref.split(':')[0], limit=8)
    _log_diffs(copier, 'insert', target_ref.split(':')[0], found)
    if before_paste is not None:
        before_paste()
    page_a.bring_to_front()
    copier._select(page_a, f'{column_name(left_col)}{target_row}')
    clip = formats.stage([key], left_col, right_col, [expected_values])
    fresh = snapshot(clip, source_ref, max_cells, strict_rows=True)
    if differences(expected, fresh, source_ref.split(':')[0], limit=1):
        raise ValueError('表A1复制值校验失败')
    copier._verify_local_clipboard(clip)
    copier._paste(page_a, target_ref, values=expected_values)
    page_a.wait_for_timeout(500)
    page_a.keyboard.press('Escape')
    _verify_rows(copier, page_a, target_ref, fresh, max_cells, target_ref)
    verify_written_formats(copier, page_a, target_ref, fresh, formats, [key], max_cells)
    return right_col - left_col + 1


def _load_mirror(copier, config):
    from local_mirror import mirror_path, STATE_SHEET
    from openpyxl import load_workbook
    import json
    path = mirror_path(copier.root, config)
    if not path.exists():
        raise ValueError('请先完成表A1镜像初始化')
    book = load_workbook(path, read_only=True)
    try:
        metadata = json.loads(book[STATE_SHEET].cell(1, 1).value)
        if not metadata.get('initialized'):
            raise ValueError('表A1镜像尚未完成初始化')
        identity = metadata['identity']
    finally:
        book.close()
    return _new_mirror(config, identity['first_col'], identity['last_col'], copier)


def run_insert(copier, config, only_rows=None):
    resources = []
    try:
        return _run_insert(copier, config, only_rows, resources)
    finally:
        for resource in reversed(resources):
            resource.close()


def _run_insert(copier, config, only_rows, resources):
    """把表A1 的新记录插入表A2。only_rows：只看这些行号（用于“改动即同步”）。"""
    pk_name, pk_number = parse_pk_column(config)
    data_start = data_start_row(config)
    max_cells = config['source']['max_cells']
    batch_rows = sync_batch_rows(config)
    max_pastes = int(config.get('sync', {}).get('max_pastes', 0) or 0)
    mirror = _load_mirror(copier, config)
    resources.append(mirror)
    formats = None
    page_c = None
    c1, c2 = mirror.first_col, mirror.last_col
    skipped = effective_skip_columns(config, c1, c2)
    page_a = _open_sheet(copier, config, 'a')
    width = c2 - c1 + 1

    cache_key = (config['upload'].get('share_url'), config['upload'].get('sheet_name'),
                 config['upload'].get('scan_range'), pk_name, data_start)
    a_index = None
    if getattr(copier, '_target_index_url', None) == cache_key:
        a_index = copier._target_index
    if a_index is None:
        # 表A标签页被关掉/崩溃时（TargetClosedError）重新打开并重试一次。
        for _attempt in range(2):
            try:
                a_index = _build_a_index(copier, config, page_a, pk_number, pk_name, data_start, skipped, c1, c2)
                break
            except Exception as exc:
                if type(exc).__name__ != 'TargetClosedError' or _attempt:
                    raise
                copier.update_upload(message='表A页面已关闭或崩溃，正在重新打开表A并重试主键扫描…')
                page_a = _open_sheet(copier, config, 'a')
        copier._target_index = a_index
        copier._target_index_url = cache_key

    # Only A1 supplies insertion values and pending primary keys.
    # Mirror rows are stable even when D reorders; use the last D scan only for
    # anchor coordinates. Values always come from A1 by primary key.
    positions = getattr(copier, '_a1_source_rows', None)
    pk_rows = sorted((row, key) for key, row in positions.items()) if positions is not None else mirror.key_rows()
    state = mirror.load_state()
    pending_keys = {key for key, value in state.items() if value['row_hash'] != value['pushed_hash']}
    inserts = [entry for entry in find_inserts(pk_rows, a_index) if entry['pk'] in pending_keys]
    if only_rows is not None:
        keys = getattr(copier, '_a1_target_keys', set())
        inserts = [entry for entry in inserts if entry['pk'] in keys]
    inserted = overwritten = header_skipped = unmappable = writes = 0
    pushed = []
    actionable = [p for p in inserts if p['target_row'] is not None]
    unmappable = len(inserts) - len(actionable)
    activity = getattr(copier, 'activity', None)
    if activity is not None:
        activity.record('开始插入新增记录', candidates=len(inserts), actionable=len(actionable))
    if max_pastes:
        # 每轮只处理最前面这么多条新增记录，写完就结束本轮，不再继续向下扫描。
        actionable = actionable[:max_pastes]
    if actionable:
        saved = mirror.read_rows([entry['pk'] for entry in actionable])
        a_rows = _read_rows_by_number(copier, page_a,
                                      sorted({p['target_row'] for p in actionable}),
                                      c1, c2, max_cells, skipped, batch_rows)
        for p in actionable:
            if _stop_requested(copier):
                raise SyncStopped()
            target_row = p['target_row']
            if target_row < data_start:
                # 锚点之上插入会落到表头：跳过，避免覆盖表头。
                header_skipped += 1
                continue
            d_values = saved[p['pk']]
            a_values = a_rows.get(target_row)
            if d_values is None or a_values is None:
                unmappable += 1
                continue
            # 目标行原本有数据也照常写入：changed_cells 只取“表C非空且与表A不同”的单元格，
            # 因此表A1的改动会覆盖表A原值（含主键列），表C为空的单元格则跳过、不清空表A。
            occupied = not _is_blank_row(a_values, c1, width, skipped)
            cells = changed_cells(d_values, a_values, c1, width, skipped)
            cells.sort(key=lambda item: item[0])
            if not cells:
                pushed.append((p['pk'], state[p['pk']]['row_hash']))
                continue
            row_written = False
            if formats is None:
                formats = copier._cell_format_source(config)
                resources.append(formats)
            groups = list(_groups(cells, lambda a, b: b[0] == a[0] + 1))
            for group in groups:
                if _stop_requested(copier):
                    raise SyncStopped()
                left_col, right_col = group[0][0], group[-1][0]
                block_values = [d_values[col_abs - c1] if col_abs - c1 < len(d_values) else ''
                                for col_abs in range(left_col, right_col + 1)]
                pasted = _paste_block(copier, page_c, page_a, p['source_row'], target_row,
                                      left_col, right_col, max_cells, block_values,
                                      before_paste=lambda: mirror.read_rows([p['pk']]),
                                      formats=formats, key=p['pk'])
                if pasted:
                    writes += pasted
                    row_written = True
            pushed.append((p['pk'], state[p['pk']]['row_hash']))
            if row_written:
                inserted += 1
                if occupied:
                    overwritten += 1

    if pushed:
        mirror.mark_pushed(pushed)
    result = dict(scanned=len(pk_rows), inserted=inserted, writes=writes,
                  overwritten=overwritten, header_skipped=header_skipped,
                  unmappable=unmappable, max_pastes=max_pastes,
                  targeted=only_rows is not None,
                  skipped_columns=format_skip_columns(skipped))
    result['message'] = (f'新增记录处理完成：扫描{len(pk_rows)}行，写入{inserted}条（{writes}格；'
                         f'其中覆盖表A原值{overwritten}条）；落到表头跳过{header_skipped}条，'
                         f'无法定位{unmappable}条。')
    if activity is not None:
        activity.record('新增记录处理完成', message=result['message'], scanned=len(pk_rows),
                        inserted=inserted, writes=writes, overwritten=overwritten,
                        header_skipped=header_skipped, unmappable=unmappable)
    copier.update(message=result['message'])
    return result
