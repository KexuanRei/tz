"""WatchDog：表A 检查者。

OverWatcher 确认表C 变化后，把变化行交给 WatchDog：按主键定位差异行，对差异单元格复制粘贴
覆盖；完成后交回 OverWatcher 继续扫描。

定时同步默认使用线下对比：表A 当前内容来自本地表B（DATA/表B.xlsx 快照），主键→行号也由表B
提供，因此不再在线读取表A 进行对比；只在粘贴后回读校验写入结果。旧检索方式（legacy_scan）
与“改动即同步”（only_rows）仍走在线读取表A 的回退路径。
写回成功后由调用方标记 表A1 镜像为已推送。
"""
from copy_range import column_name
from upload_verify import differences, snapshot
from cell_paste import changed_spans, plan_groups, verify_written_formats
from sync_engine import pk_key, sync_batch_rows
from sync_columns import format_skip_columns


class WatchDog:
    """表A 检查者：按主键定位差异行并覆盖差异单元格。"""

    def __init__(self, copier, config, mirror, page_c, page_a, c1, c2, pk_name, pk_number,
                 pk_offset, width, spans, skipped, max_cells, max_pastes,
                 a_map, rebuild_map, cache_key, table_b=None, table_b_merges=None):
        self.copier = copier
        self.config = config
        self.mirror = mirror
        self.page_c = page_c
        self.page_a = page_a
        self.c1 = c1
        self.c2 = c2
        self.pk_name = pk_name
        self.pk_number = pk_number
        self.pk_offset = pk_offset
        self.width = width
        self.spans = spans
        self.skipped = skipped
        self.max_cells = max_cells
        self.max_pastes = max_pastes
        self.a_map = a_map
        self.rebuild_map = rebuild_map
        self.cache_key = cache_key
        # 表B：线上表A 的本地 xlsx 快照 {主键: (表A行号, 值列表)}。提供时，
        # 表A 当前内容与主键→行号映射都从本地读取，不再在线读取表A。
        self.table_b = table_b
        self.table_b_merges = set(table_b_merges or ())
        if table_b is not None:
            self.a_map = {key: row for key, (row, _values) in table_b.items()}
        self.changed = 0
        self.missing = 0
        self.paste_batches = 0
        self.stopped_early = False
        self.formats = None

    def close(self):
        if self.formats is not None:
            self.formats.close()
            self.formats = None

    def _ref(self, first, last):
        return f'{column_name(self.c1)}{first}:{column_name(self.c2)}{last}'

    def _rebuild_index(self):
        self.a_map = self.rebuild_map()
        self.copier._target_index = self.a_map
        self.copier._target_index_url = self.cache_key

    def _merge_overlaps(self, left, top, right, bottom):
        """表B 快照里是否有合并单元格与目标矩形区域重叠。"""
        for (mc1, mr1, mc2, mr2) in self.table_b_merges:
            if mc2 >= left and mc1 <= right and mr2 >= top and mr1 <= bottom:
                return True
        return False

    def process_entries(self, entries):
        """处理 OverWatcher 确认的表C 变化行：定位表A 差异行并覆盖差异单元格。"""
        from sync_runner import (_groups, _read_kept_columns,
                                 _read_block, _verify_rows,
                                 _record, _log_diffs, _stop_requested, SyncStopped,
                                 _read_row_windows)
        copier, page_c, page_a = self.copier, self.page_c, self.page_a
        c1 = self.c1
        pk_name, pk_offset, pk_number, width = self.pk_name, self.pk_offset, self.pk_number, self.width
        spans = self.spans
        skipped, max_cells, max_pastes = self.skipped, self.max_cells, self.max_pastes

        # 读取允许跨过非候选行，按矩形跨度分批；粘贴仍严格要求连续。
        if self.table_b is not None:
            # 本地对比：表A 当前内容与主键→行号都来自表B，不再在线读取表A。
            matched = [dict(entry, target=self.table_b[entry['pk']][0])
                       for entry in entries if entry['pk'] in self.table_b]
            actual_rows = {entry['pk']: self.table_b[entry['pk']][1] for entry in matched}
        else:
            for attempt in range(2):
                matched = [dict(entry, target=self.a_map[entry['pk']])
                           for entry in entries if entry['pk'] in self.a_map]
                actual_rows = {}
                moved = False
                by_target = {entry['target']: entry for entry in matched}
                for row_numbers in _read_row_windows(by_target, sync_batch_rows(self.config), max_cells, width):
                    if _stop_requested(copier):
                        raise SyncStopped()
                    first_row, last_row = row_numbers[0], row_numbers[-1]
                    target_ref = self._ref(first_row, last_row)
                    copier.update(message=f'正在批量检查表A {target_ref}（{len(row_numbers)}条候选记录）…')
                    actual = _read_kept_columns(copier, page_a, target_ref, max_cells, skipped)
                    for row_no in row_numbers:
                        entry = by_target[row_no]
                        row = actual['rows'][row_no - first_row]
                        if pk_key(row[pk_offset]) != entry['pk']:
                            moved = True
                            break
                        actual_rows[entry['pk']] = row
                    if moved:
                        break
                if not moved:
                    break
                if attempt:
                    raise ValueError('表A行位置持续变化，已停止本轮，请勿在同步时排序或插入行')
                self._rebuild_index()

        entry_by_source = {entry['source']: entry for entry in entries}
        for entry in entries:
            if entry['pk'] not in self.a_map:
                self.missing += 1
                _record(copier, pk=entry['pk'], source_row=entry['source'], target_row=None,
                        status='missing', next_source_row=entry['source'] + 1)

        pending = []
        synced = set()
        for entry in matched:
            current = actual_rows[entry['pk']]
            if changed_spans(entry['row'], current, c1, skipped):
                pending.append(entry)
            else:
                synced.add(entry['source'])
        if max_pastes:
            remaining = max_pastes - self.changed
            if remaining <= 0:
                self.stopped_early = True
                return
            pending = pending[:remaining]

        write_groups = plan_groups(pending, actual_rows, c1, skipped)
        todo = {entry['pk']: set() for entry in pending}
        for group, left, right in write_groups:
            for entry in group:
                todo[entry['pk']].update(range(left, right + 1))
        if write_groups:
            if self.formats is None:
                self.formats = copier._cell_format_source(self.config)
            # Validate every planned C key before the first write in this batch.
            for group, left, right in write_groups:
                self.formats.preflight([e['pk'] for e in group], left, right)
        if write_groups and page_a is None:
            # 本地对比确实有差异单元格要写回时才惰性打开表A；本地无差异则不开表A，
            # 直接标记已同步进入下一轮。
            from sync_runner import _open_sheet
            page_a = _open_sheet(copier, self.config, 'a')
            self.page_a = page_a
        written_rows = set()
        write_groups.reverse()
        while write_groups:
            if _stop_requested(copier):
                raise SyncStopped()
            group, left_col, right_col = write_groups.pop()
            first, last = group[0], group[-1]
            source_ref = f'{column_name(left_col)}{first["source"]}:{column_name(right_col)}{last["source"]}'
            target_ref = f'{column_name(left_col)}{first["target"]}:{column_name(right_col)}{last["target"]}'
            keys = [e['pk'] for e in group]
            progress = dict(pk=first['pk'], primary_keys=keys,
                            source_row=first['source'], source_end=last['source'],
                            target_row=first['target'], target_end=last['target'])
            _record(copier, **progress, status='pending', next_source_row=first['source'])
            copier.update_upload(message=f'正在粘贴差异单元格 {target_ref}，按表C同主键同列应用格式…')
            expected = dict(rows=[e['row'][left_col-c1:right_col-c1+1] for e in group],
                            row_count=len(group), width=right_col-left_col+1)
            if self.table_b is not None:
                # 线下对比：表A 当前内容来自表B，不再在线读取表A 做二次比对或定位。
                if self._merge_overlaps(left_col, first['target'], right_col, last['target']):
                    raise ValueError(f'表A {target_ref} 涉及合并单元格，无法只更新差异单元格')
                offline_actual = dict(rows=[actual_rows[e['pk']][left_col-c1:right_col-c1+1] for e in group],
                                      row_count=len(group), width=right_col-left_col+1)
                found = differences(expected, offline_actual, target_ref.split(':')[0], limit=8)
                _log_diffs(copier, 'paste', target_ref.split(':')[0], found)
            else:
                target = _read_block(copier, page_a, target_ref, max_cells)
                if target['merges']:
                    raise ValueError(f'表A {target_ref} 涉及合并单元格，无法只更新差异单元格')
                target_pk = _read_block(copier, page_a,
                    f'{pk_name}{first["target"]}:{pk_name}{last["target"]}', max_cells)
                if [pk_key(row[0]) for row in target_pk['rows']] != keys:
                    raise ValueError('表A行位置在写入前发生变化，已停止本轮，下轮重新按主键定位')
                # A collaborator may have fixed some cells since the initial read.
                # Shrink again; never let an unchanged cell remain inside a paste.
                current = {e['pk']: row for e, row in zip(group, target['rows'])}
                narrowed = [dict(e, row=row) for e, row in zip(group, expected['rows'])]
                parts = plan_groups(narrowed, current, left_col)
                selected = {e['pk']: set() for e in group}
                for part, left, right in parts:
                    for entry in part:
                        selected[entry['pk']].update(range(left, right + 1))
                for entry in group:
                    todo[entry['pk']].difference_update(set(range(left_col, right_col+1)) - selected[entry['pk']])
                same = (len(parts) == 1 and parts[0][1:] == (left_col, right_col)
                        and [e['pk'] for e in parts[0][0]] == keys)
                if not same:
                    originals = {e['pk']: e for e in group}
                    write_groups.extend(([originals[e['pk']] for e in part], left, right)
                                        for part, left, right in reversed(parts))
                    continue
                found = differences(expected, target, target_ref.split(':')[0], limit=8)
                _log_diffs(copier, 'paste', target_ref.split(':')[0], found)
            saved = self.mirror.read_rows(keys)
            rows = [saved[e['pk']][left_col-c1:right_col-c1+1] for e in group]
            actual_saved = dict(rows=rows, row_count=len(rows), width=right_col-left_col+1)
            if differences(expected, actual_saved, source_ref.split(':')[0], limit=1):
                raise ValueError('表A1内容在写入前发生变化，已停止本轮')
            page_a.bring_to_front()
            copier._select(page_a, f'{column_name(left_col)}{first["target"]}')
            clip = self.formats.stage(keys, left_col, right_col, rows)
            fresh = snapshot(clip, source_ref, max_cells, strict_rows=True)
            if differences(expected, fresh, source_ref.split(':')[0], limit=1):
                raise ValueError('表A1带格式复制值校验失败，尚未粘贴表A')
            copier._verify_local_clipboard(clip)
            copier._paste(page_a, target_ref, values=rows)
            page_a.wait_for_timeout(500)
            # Dismiss WPS's paste-options button/menu before the next Ctrl+C.
            # Choosing "粘贴值" here would discard the requested table-C styles.
            page_a.keyboard.press('Escape')
            _verify_rows(copier, page_a, target_ref, fresh, max_cells, target_ref)
            # This is a fresh, complete read-back for format verification, not
            # the last partial clipboard left by a split value read.
            verify_written_formats(copier, page_a, target_ref, fresh, self.formats, keys, max_cells)
            for entry in group:
                written_rows.add(entry['source'])
                todo[entry['pk']].difference_update(range(left_col, right_col + 1))
            self.paste_batches += 1
            complete = all(not todo[e['pk']] for e in group)
            _record(copier, **progress, status='verified' if complete else 'partial',
                    next_source_row=last['source'] + 1 if complete else first['source'],
                    source_range=source_ref, target_range=target_ref,
                    skipped_columns=format_skip_columns(skipped))
            activity = getattr(copier, 'activity', None)
            if activity is not None:
                activity.record('写回表A完成', source=source_ref, target=target_ref,
                                rows=len(group), columns=f'{column_name(left_col)}～{column_name(right_col)}',
                                cells=(right_col - left_col + 1) * len(group),
                                complete=complete)

        for entry in pending:
            if not todo[entry['pk']]:
                synced.add(entry['source'])
        self.changed += len(written_rows)
        pushed = [(entry_by_source[s]['pk'], entry_by_source[s]['hash']) for s in synced]
        if pushed:
            self.mirror.mark_pushed(pushed)
            activity = getattr(copier, 'activity', None)
            if activity is not None:
                activity.record('表A1标记已推送', rows=len(pushed))
        activity = getattr(copier, 'activity', None)
        if self.missing and activity is not None:
            activity.record('表A缺少主键', rows=self.missing)
        if max_pastes and self.changed >= max_pastes:
            self.stopped_early = True
