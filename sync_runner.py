"""按自定义批大小读取表C，按主键合并复制连续差异行，再继续下一批。

新流程：表C → 本地 表A1 镜像（按主键/分段部做变更检测）→ 表A。
- 每一批先把表C 读进内存，与 表A1 镜像比对：内容有变化就更新镜像，尚未写回表A 的行进入写回候选。
- 只对候选行做表A 定位、比对和粘贴；写回成功后把 表A1 里对应行标记为“已推送”。
"""
import json
from pathlib import Path
import time
import urllib.parse

from copy_range import column_name
from update_excel import address
from excel_clipboard import inspect_copy
from clipboard_html import CopyIntegrityError
from upload_verify import snapshot, differences, mismatch
from sync_engine import (parse_pk_column, data_start_row, sync_batch_rows, build_index,
                         pk_key, MAX_BATCH_ROWS, row_hash, start_cell, legacy_scan)
from watch import install as install_watch
from sync_columns import (sync_skip_columns, kept_spans, format_skip_columns,
                          effective_skip_columns, parse_sync_columns)
from watchdog import WatchDog
from local_mirror import LocalMirror, mirror_values

# 同一处不一致持续这么久才判定失败（WPS 粘贴后异步渲染）。
MISMATCH_SETTLE_SECONDS = 8.0
READBACK_TIMEOUT_SECONDS = 30.0

# 网页剪贴板富文本有大小上限：一次性复制整张几十万格的大表时，富文本会被截断成
# 几千行（纯文本却仍完整），无法验证完整性。因此整表复制按行分批、每批整列宽复制，
# 每批行数与格数同时封顶，确保单次复制远小于网页截断点，也绝不按列拆分重试。
_WHOLE_COPY_MAX_ROWS = 1500      # 单次复制最多行数（实测 76 列时约 2300+ 行即被截断）
_WHOLE_COPY_MAX_CELLS = 100_000  # 单次复制最多格数


class SyncStopped(Exception):
    """用户主动请求停止当前同步轮次；由 EdgeCopier 捕获并优雅结束，不算错误。"""
    code = 'SYNC_STOPPED'


def _elapsed_text(start):
    """把起始 monotonic 时间换算成「秒」文本，用于操作日志的耗时字段。"""
    return f'{time.monotonic() - start:.3f}s'


def _stop_requested(copier):
    """测试用的 FakeCopier 没有 should_stop，此处防御性判断。"""
    check = getattr(copier, 'should_stop', None)
    return bool(check and check())


def read_online_grid(copier, page, config, ref, label):
    """把在线表读成二维网格（rows 为按表头行排序的值列表，header 为表头）。"""
    cfg = config['source']
    if ref:
        left, right = ref.split(':')
        c1, r1 = address(left)
        c2, r2 = address(right)
        if c2 < c1 or r2 < r1:
            raise ValueError(f'{label}扫描区域无效')
        batch_rows = min(max(10, int(cfg.get('batch_rows') or 1000)), MAX_BATCH_ROWS)
        width = c2 - c1 + 1
        rows, header = [], None
        cursor = r1
        while cursor <= r2:
            count = min(batch_rows, r2 - cursor + 1)
            while True:
                end = cursor + count - 1
                bref = f'{column_name(c1)}{cursor}:{column_name(c2)}{end}'
                copier.update(message=f'正在扫描{label} {bref}…')
                try:
                    clip = copier._capture(page, bref)
                    grid = inspect_copy(clip, config, f'{column_name(c1)}{cursor}', 'selection',
                                        expected_range=bref, batch=True, strict_rows=True)
                except CopyIntegrityError as exc:
                    if exc.can_split and count > 1:
                        count = max(1, count // 2)
                        continue
                    raise
                if header is None:
                    header = grid['header']
                rows.extend(grid['rows'])
                cursor += count
                break
        return dict(rows=rows, header=header, width=width,
                    first_column=c1, first_row=r1, row_count=len(rows))
    # 整表一次读取（适合小表；大表请配置扫描区域）。
    copier.update(message=f'正在扫描{label}整张工作表…')
    clip = copier._capture(page, None)
    grid = inspect_copy(clip, config, 'A1', 'all')
    return dict(rows=grid['rows'], header=grid['header'], width=grid['width'],
                first_column=grid['first_column'], first_row=grid['first_row'],
                row_count=grid['row_count'])


def _verify_rows(copier, page, ref, expected, max_cells, label):
    """粘贴后持续读回直到稳定；不一致才报错（WPS 异步渲染）。"""
    origin = ref.split(':')[0]
    start = time.monotonic()
    last = None
    first_seen = None
    samples = 0
    while True:
        now = time.monotonic()
        copier.update_upload(message=f'正在回读校验表A {ref}（仅检查，不重复粘贴）…')
        actual = _read_snapshot_split(copier, page, ref, max_cells)
        found = differences(expected, actual, origin)
        if not found:
            return actual
        diff = mismatch(expected, actual, origin) if found else None
        if diff != last:
            last, first_seen, samples = diff, now, 1
        else:
            samples += 1
        stable = now - first_seen if first_seen is not None else 0.0
        if (samples >= 2 and stable >= MISMATCH_SETTLE_SECONDS) or now - start >= READBACK_TIMEOUT_SECONDS:
            raise ValueError(f'表A {label} 回读校验失败：{diff}')
        page.wait_for_timeout(min(400 + 400 * samples, 1500))


def _open_sheet(copier, config, which):
    from online_format import OnlineFormatter
    section, label = ('source', '表C') if which == 'c' else ('upload', '表A')
    opener = copier._page_c if which == 'c' else copier._page_a
    page = opener(config[section]['share_url'], reload=True)
    page.bring_to_front()
    install_watch(page)
    url = urllib.parse.urlparse(page.url)
    if url.scheme != 'https' or url.hostname not in {'www.kdocs.cn', 'kdocs.cn'}:
        raise ValueError(f'请先打开{label}并登录，完成后再同步')
    copier.context.grant_permissions(['clipboard-read', 'clipboard-write'],
                                     origin=url.scheme + '://' + url.netloc)
    formatter = OnlineFormatter(page, lambda ref: copier._select(page, ref),
                                progress=lambda text: copier.update(message=text.replace('表C', label)))
    formatter.select_sheet(config[section]['sheet_name'])
    activity = getattr(copier, 'activity', None)
    if activity is not None:
        activity.record(f'打开{label}工作表', sheet=config[section]['sheet_name'])
    return page


def _bounds(copier, page, config):
    ref = config['source'].get('scan_range')
    if not ref and (sync_skip_columns(config)
                    or parse_sync_columns(config.get('sync', {}).get('sync_columns', ''))):
        ref = copier._used_range(page)
    if ref:
        left, right = ref.split(':')
        c1, r1 = address(left)
        c2, r2 = address(right)
    else:
        # 未指定范围时，仅用整表复制识别边界；实际比较仍逐行重新读取。
        grid = read_online_grid(copier, page, config, None, '表C（识别范围）')
        c1, r1 = grid['first_column'], grid['first_row']
        c2, r2 = c1 + grid['width'] - 1, r1 + grid['row_count'] - 1
    if c1 < 1 or c2 < c1 or r2 < r1:
        raise ValueError('表C扫描范围为空或无效')
    start = max(r1, data_start_row(config))
    return c1, start, c2, r2


def _read_snapshot_split(copier, page, ref, max_cells, pad_trailing_columns=False):
    """复制缺行时按行拆分，缺列时按列拆分；每个子块独立验证后按原坐标合并。

    不能把缺少的一列直接认作末尾空白列：也可能是中间隐藏列。
    此返回值仅供比较；粘贴仍必须使用一次完整复制的真实剪贴板。
    """
    left, right = ref.split(':')
    c1, r1 = address(left)
    c2, r2 = address(right)
    from read_limits import check_cells
    check_cells(r2 - r1 + 1, c2 - c1 + 1, max_cells, '同步读取范围')
    todo = [(c1, r1, c2, r2)]
    chunks = []
    while todo:
        lc, lo, rc, hi = todo.pop()
        sub = f'{column_name(lc)}{lo}:{column_name(rc)}{hi}'
        try:
            result = snapshot(copier._capture(page, sub), sub, max_cells, strict_rows=True,
                              pad_trailing_columns=pad_trailing_columns)
        except CopyIntegrityError as exc:
            if exc.can_split and exc.kind == 'columns' and rc > lc:
                mid = (lc + rc) // 2
                copier.update(message=f'{sub} 复制缺列，正在按列拆分读取并核对原列位置…')
                todo.append((mid + 1, lo, rc, hi))
                todo.append((lc, lo, mid, hi))
                continue
            if exc.can_split and exc.kind != 'columns' and hi > lo:
                mid = (lo + hi) // 2
                todo.append((lc, mid + 1, rc, hi))
                todo.append((lc, lo, rc, mid))
                continue
            raise
        chunks.append((lc, lo, result))
    rows = [[''] * (c2 - c1 + 1) for _ in range(r2 - r1 + 1)]
    merges = set()
    unverified = 0
    for _lc, _lo, result in chunks:
        for offset, values in enumerate(result['rows']):
            rows[_lo - r1 + offset][_lc - c1:_lc - c1 + result['width']] = values
        merges.update((left + _lc - c1, top + _lo - r1,
                       right + _lc - c1, bottom + _lo - r1)
                      for left, top, right, bottom in result['merges'])
        unverified += int(result.get('unverified_rows') or 0)
    return dict(rows=rows, row_count=len(rows), width=c2 - c1 + 1,
                merges=merges, unverified_rows=unverified)


def _read_block(copier, page, ref, max_cells, pad_trailing_columns=False):
    result = _read_snapshot_split(copier, page, ref, max_cells,
                                  pad_trailing_columns=pad_trailing_columns)
    if any(bottom != top for _left, top, _right, bottom in result['merges']):
        raise ValueError(f'{ref} 含跨行合并单元格，无法安全按主键分批复制')
    return result


def _read_kept_columns(copier, page, ref, max_cells, skipped):
    """Read only retained column spans; explicit exclusions get neutral comparison slots."""
    c1, r1 = address(ref.split(':')[0])
    c2, r2 = address(ref.split(':')[1])
    spans = list(kept_spans(c1, c2, skipped))
    if spans == [(c1, c2)]:
        return _read_block(copier, page, ref, max_cells)
    from read_limits import check_cells
    check_cells(r2-r1+1, c2-c1+1, max_cells, '同步读取范围')
    rows = [['']*(c2-c1+1) for _ in range(r2-r1+1)]
    merges = set()
    for left, right in spans:
        # 连续保留区间内没有跳过/隐藏列，缺失的末尾列必为空白，可安全补齐，
        # 避免空白单元格把整段读取拆成一格一格。
        block = _read_block(copier, page, f'{column_name(left)}{r1}:{column_name(right)}{r2}', max_cells,
                            pad_trailing_columns=True)
        for offset, values in enumerate(block['rows']):
            rows[offset][left-c1:right-c1+1] = values
        merges.update((a+left-c1,b,c+left-c1,d) for a,b,c,d in block['merges'])
    return dict(rows=rows, row_count=len(rows), width=c2-c1+1, merges=merges)


def _whole_data_ref(copier, page, config, label):
    """从起始单元格到最后一个有数据的单元格，算出整表复制区域。

    模拟人类操作：起始单元格 + Ctrl+End 定位表尾（复用 _used_range），
    返回 (起始列, 起始行, 结束列, 结束行, 区域 ref)。
    """
    col_name, start_col, start_row = start_cell(config)
    used = copier._used_range(page)
    last_col, last_row = address(used.split(':')[1])
    if last_row < start_row or last_col < start_col:
        raise ValueError(f'{label}没有数据：起始单元格 {col_name}{start_row} 之后没有有数据的行')
    ref = f'{col_name}{start_row}:{column_name(last_col)}{last_row}'
    return start_col, start_row, last_col, last_row, ref


def _whole_copy_rows_per_copy(widest):
    """整表复制时每批最多复制的行数：按最宽保留区间的格数封顶，行数再封顶，绝不按列拆分。"""
    return max(1, min(_WHOLE_COPY_MAX_ROWS, _WHOLE_COPY_MAX_CELLS // max(1, widest)))


def read_whole_table(copier, page, config, label):
    """整表扫描（起始单元格 → 表尾），返回与批量读取同构的网格。

    模拟人类操作：起始单元格 + Ctrl+End 定位表尾，按“跳过列”切出的保留区间复制
    （跳过列/隐藏列不复制）。网页剪贴板富文本有大小上限，一次性复制整张几十万格的
    大表时富文本会被截断成几千行（纯文本却仍完整），无法验证完整性——这正是此前
    “复制缺列/缺行”反复循环、长时间卡住的原因。因此按行分批、每批整列宽复制，
    末尾空白行/列按空白补齐，不再按行/列二分重试。
    """
    start_col, start_row, last_col, last_row, ref = _whole_data_ref(copier, page, config, label)
    max_cells = config['source']['max_cells']
    skipped = effective_skip_columns(config, start_col, last_col)
    copier.update(message=f'正在整表复制{label} {ref}…')
    width = last_col - start_col + 1
    total_rows = last_row - start_row + 1
    rows = [[''] * width for _ in range(total_rows)]
    merges = set()
    spans = list(kept_spans(start_col, last_col, skipped))
    widest = max((right - left + 1) for left, right in spans) if spans else width
    rows_per_copy = _whole_copy_rows_per_copy(widest)
    for batch_start, batch_end in _row_batches(start_row, last_row, rows_per_copy):
        if _stop_requested(copier):
            raise SyncStopped()
        copier.update(message=f'正在整表复制{label} 第{batch_start}～{batch_end}行…',
                      scan_row=batch_start, scan_end=last_row, batch_end=batch_end)
        idx0 = batch_start - start_row
        for left, right in spans:
            span_ref = f'{column_name(left)}{batch_start}:{column_name(right)}{batch_end}'
            block = snapshot(copier._capture(page, span_ref), span_ref, max_cells,
                             strict_rows=False, pad_trailing_columns=True, pad_trailing_rows=True)
            for offset, values in enumerate(block['rows']):
                if idx0 + offset >= total_rows:
                    break
                rows[idx0 + offset][left - start_col:right - start_col + 1] = values
            merges.update((a + left - start_col, b + idx0, c + left - start_col, d + idx0)
                          for a, b, c, d in block['merges'])
    return dict(rows=rows, first_row=start_row, first_col=start_col,
                width=width, row_count=total_rows, last_row=last_row, last_col=last_col,
                merges=merges)


def read_source_file(copier, page, config, label):
    """方法2：下载表C 原始 xlsx（复用“下载表C”机制），从文件读取整表数据。

    网页剪贴板富文本有大小上限，超大表一次复制会被截断；下载整表文件则不受此限。
    每轮“sync”任务开始时清空 copier._source_snapshot，本轮内 build_mirror 与
    run_sync 共享同一次下载与读取，避免重复下载。
    """
    cached = getattr(copier, '_source_snapshot', None)
    if cached is not None:
        return cached
    from table_c_download import download_table_c
    import source_file
    activity = getattr(copier, 'activity', None)
    if not getattr(copier, '_table_c_fresh', False):
        # 启动时已下载过表C 则直接复用，避免同一轮重复下载两次。
        if activity is not None:
            activity.record('开始下载表C', path='DATA/表C.xlsx')
        copier.update(message=f'正在下载{label}整表文件…')
        start = time.monotonic()
        result = download_table_c(copier, config, page=page)
        if activity is not None:
            activity.record('表C下载完成', path=result.get('path'), sheet=result.get('sheet'),
                            size=result.get('size'), 耗时=_elapsed_text(start))
        copier.update(message=f'正在读取{label}整表数据…')
        # 回填格式参考路径：方法2下载的表C原始文件同时作为表C格式参考。
        if result and result.get('path'):
            config.setdefault('source', {})['format_file'] = result['path']
            if result.get('sheet'):
                config['source']['format_sheet'] = result['sheet']
            writer = getattr(copier, '_write_config_file', None)
            if writer is not None:
                try:
                    writer(config)
                except Exception:
                    pass
    else:
        copier._table_c_fresh = False
        copier.update(message=f'正在读取{label}整表数据…')
    read_start = time.monotonic()
    grid = source_file.read_table_data(copier.root, config)
    if activity is not None:
        activity.record('读取表C本地文件', path=config.get('source', {}).get('format_file') or 'DATA/表C.xlsx',
                        sheet=config.get('source', {}).get('format_sheet') or '',
                        rows=grid.get('row_count'), cols=grid.get('width'),
                        merges=len(grid.get('merges') or ()), 耗时=_elapsed_text(read_start))
    copier._source_snapshot = grid
    return grid


def read_table_b_source(copier, config, c1, c2):
    """下载表A 为本地表B，并读取成 {主键: (表A行号, 值)} 供本地对比。

    方法与表C 完全一致：每轮“sync”开始时清空 copier._table_b_snapshot，
    本轮内只下载/读取一次；下载完成即读取，避免重复下载表A。
    """
    cached = getattr(copier, '_table_b_snapshot', None)
    if cached is not None:
        return cached
    from table_c_download import download_table_b
    from table_b_compare import read_table_b, table_b_merges
    activity = getattr(copier, 'activity', None)
    if activity is not None:
        activity.record('开始下载表B', path='DATA/表B.xlsx')
    copier.update_upload(message='正在下载表A整表文件为表B…')
    start = time.monotonic()
    result = download_table_b(copier, config)
    if activity is not None:
        activity.record('表B下载完成', path=result.get('path'), sheet=result.get('sheet'),
                        size=result.get('size'), 耗时=_elapsed_text(start))
    copier.update_upload(message='正在读取表B整表数据…')
    read_start = time.monotonic()
    snapshot = read_table_b(copier.root, config, c1, c2)
    if activity is not None:
        activity.record('读取表B本地文件', path='DATA/表B.xlsx',
                        rows=len(snapshot), 耗时=_elapsed_text(read_start))
    copier._table_b_snapshot = snapshot
    # 表B 的合并单元格用于线下判断粘贴目标区域是否安全（不再在线读取表A）。
    copier._table_b_merges = table_b_merges(copier.root, config)
    if activity is not None:
        activity.record('读取表B合并单元格', path='DATA/表B.xlsx',
                        merges=len(copier._table_b_merges))
    return snapshot


def read_pk_column_index(copier, page, config, pk_number, pk_name, label, start_row=None):
    """新方式建索引：直接选中主键列（起始行 → 最后有数据行）复制，返回 {主键: 行号}。

    主键列只有一列，但网页富文本同样有大小上限；超高列按行分批复制，避免截断。
    """
    max_cells = config['source']['max_cells']
    if start_row is None:
        start_row = data_start_row(config)
    start = time.monotonic()
    used = copier._used_range(page)
    last_col, last_row = address(used.split(':')[1])
    if last_row < start_row:
        return {}
    copier.update_upload(message=f'正在整列复制{label}主键 {pk_name}…')
    result = {}
    rows_per_copy = _whole_copy_rows_per_copy(1)
    for batch_start, batch_end in _row_batches(start_row, last_row, rows_per_copy):
        if _stop_requested(copier):
            raise SyncStopped()
        ref = f'{column_name(pk_number)}{batch_start}:{column_name(pk_number)}{batch_end}'
        block = snapshot(copier._capture(page, ref), ref, max_cells,
                         strict_rows=False, pad_trailing_columns=True, pad_trailing_rows=True)
        for offset, row in enumerate(block['rows']):
            sheet_row = batch_start + offset
            key = pk_key(row[0] if row else '')
            if key is None:
                continue
            if key in result:
                raise ValueError(
                    f'{label}主键列 {pk_name} 存在重复值“{key}”（第{result[key]}行与第{sheet_row}行），'
                    f'无法唯一匹配。请更换主键列，或修正{label}中的重复值。')
            result[key] = sheet_row
    activity = getattr(copier, 'activity', None)
    if activity is not None:
        activity.record('读取主键列', label=label, column=pk_name, rows=len(result),
                        range=f'{pk_name}{start_row}:{pk_name}{last_row}',
                        耗时=_elapsed_text(start))
    return result


def _record(copier, **values):
    """发现差异后立刻持久化主键和游标；不记录整行内容。"""
    path = Path(copier.root) / 'sync_progress.json'
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding='utf-8')
    temp.replace(path)


def _log_diffs(copier, label, origin, found):
    """把“判定为差异”的单元格写入 sync_diff_log.jsonl，排查“看似无差异却粘贴”。

    每个差异记录：单元格地址、表C/表A的原始值 repr 与归一化后 repr。
    通过原始值即可看出是哪种字符或格式差异（例如窄空格、日期格式、数值格式）。
    写日志失败不影响同步本身。
    """
    try:
        from upload_verify import normalize_text, comparison_result
        path = Path(copier.root) / 'sync_diff_log.jsonl'
        details = []
        for item in list(found)[:50]:
            if item.get('kind') == 'shape':
                details.append(dict(kind='shape', expected=item.get('expected'),
                                    actual=item.get('actual')))
                continue
            expected = item.get('expected')
            actual = item.get('actual')
            details.append(dict(cell=item.get('cell'),
                                d_raw=repr(expected), a_raw=repr(actual),
                                d_norm=repr(normalize_text(expected)),
                                a_norm=repr(normalize_text(actual)),
                                d_result=str(comparison_result(expected)),
                                a_result=str(comparison_result(actual))))
        line = json.dumps(dict(label=label, origin=origin, count=len(found),
                               diffs=details), ensure_ascii=False)
        with path.open('a', encoding='utf-8') as stream:
            stream.write(line + '\n')
    except Exception:
        pass


def _groups(items, consecutive):
    """把有序条目拆成可连续读取或粘贴的矩形组。"""
    group = []
    for item in items:
        if group and not consecutive(group[-1], item):
            yield group
            group = []
        group.append(item)
    if group:
        yield group


def _read_row_windows(rows, batch_rows, max_cells, width=1):
    """把离散行合并为有界读取窗口；允许跨空隙读取，但调用方只处理候选行。

    限制的是矩形实际跨度，不是候选行数量，避免远距离行合成超大范围。
    此分组只用于读取，不能用于粘贴。
    """
    from read_limits import MAX_CELLS, check_cells
    check_cells(1, width, max_cells, '同步读取范围')
    limit = max(1, min(batch_rows, MAX_BATCH_ROWS, min(max_cells, MAX_CELLS) // width))
    group = []
    for row in sorted(set(rows)):
        if group and row - group[0] >= limit:
            yield group
            group = []
        group.append(row)
    if group:
        yield group


def _row_batches(start, end, batch_rows, only_rows=None):
    """把要处理的行拆成 (first,last) 区间；only_rows 为 None 时按批全表扫描。"""
    if only_rows is None:
        cursor = start
        while cursor <= end:
            yield cursor, min(end, cursor + batch_rows - 1)
            cursor += batch_rows
        return
    rows = sorted(r for r in set(only_rows) if start <= r <= end)
    if not rows:
        return
    first = prev = rows[0]
    for r in rows[1:]:
        if r <= prev + 1 and r - first < batch_rows:
            prev = r
            continue
        yield first, prev
        first = prev = r
    yield first, prev


def _new_mirror(config, first_col, last_col, copier=None):
    root = getattr(copier, 'root', Path(__file__).resolve().parent)
    cached = getattr(copier, '_a1_mirror', None)
    from local_mirror import mirror_path
    identity = (str(mirror_path(root, config)), config.get('local', {}).get('sheet_name', '表A1'),
                first_col, last_col, json.dumps(config['source'], sort_keys=True),
                config.get('sync', {}).get('primary_key_column', 'D'),
                config.get('sync', {}).get('data_start_row', 2),
                tuple(effective_skip_columns(config, first_col, last_col)))
    if (cached is not None and getattr(copier, '_a1_mirror_key', None) == identity
            and not cached.dirty and cached.unchanged_on_disk()):
        return cached
    mirror = LocalMirror(root, config, first_col, last_col,
                         getattr(copier, 'activity', None) if copier is not None else None)
    if copier is not None:
        copier._a1_mirror = mirror
        copier._a1_mirror_key = identity
    return mirror


def _mirror_marker_path(copier):
    return Path(getattr(copier, 'root', Path(__file__).resolve().parent)) / 'mirror_state.json'


def _mirror_signature(copier, config):
    """镜像位置签名 = 文件绝对路径 + 工作表名。任一变都视为“换了位置”。"""
    from local_mirror import mirror_path
    path = mirror_path(getattr(copier, 'root', Path(__file__).resolve().parent), config)
    sheet = str(config.get('local', {}).get('sheet_name') or '表A1').strip()
    return str(path), f'{path}|{sheet}'


def _previous_mirror_signature(copier):
    try:
        data = json.loads(_mirror_marker_path(copier).read_text(encoding='utf-8'))
        return data.get('mirror')
    except Exception:
        return None


def _save_mirror_signature(copier, signature):
    marker = _mirror_marker_path(copier)
    temp = marker.with_suffix('.tmp')
    temp.write_text(json.dumps({'mirror': signature}, ensure_ascii=False, indent=2), encoding='utf-8')
    temp.replace(marker)


def _reset_mirror_at(path):
    """删除指定位置的旧镜像（仅限本程序创建的镜像），以便强制重建。

    不是本程序创建的镜像（无 _sync_state 表）不删除，交由 LocalMirror 就地“收养”
    为新镜像，避免误删用户自己的 Excel 文件。
    """
    from openpyxl import load_workbook
    from local_mirror import STATE_SHEET
    is_mirror = False
    try:
        book = load_workbook(path, read_only=True)
        try:
            is_mirror = STATE_SHEET in book
        finally:
            book.close()
    except Exception:
        return  # 无法读取（损坏/加密/占用）：交给 LocalMirror 处理
    if is_mirror:
        try:
            path.unlink(missing_ok=True)
        except PermissionError:
            raise ValueError('镜像位置已更改，但新位置上的镜像文件正被 Excel/WPS 占用，请先关闭该文件后重试') from None


def build_mirror(copier, config):
    """OverWatcher 第一阶段：整体扫描表C，把全部行上传到本地 表A1 镜像，不写表A。

    这是“构建镜像”阶段——只建立表C 的内容基线（row_hash），初始行同时标记为基线完成。之后由守护扫描（run_sync）按需把变化行写回表A。构建阶段不打开表A，
    因此即使表A 写入很慢或暂时失败，镜像也能先建好，后续按 pushed_hash 断点续跑。
    """
    if _stop_requested(copier):
        raise SyncStopped()
    pk_name, pk_number = parse_pk_column(config)
    batch_rows = sync_batch_rows(config)
    max_cells = config['source']['max_cells']
    page_c = _open_sheet(copier, config, 'c')
    whole = None
    if legacy_scan(config):
        c1, start, c2, end = _bounds(copier, page_c, config)
    else:
        # 方法2：下载表C原始文件后从文件读取整表，导入镜像（绕开剪贴板截断）。
        whole = read_source_file(copier, page_c, config, '表C')
        c1, start = whole['first_col'], whole['first_row']
        c2 = whole['first_col'] + whole['width'] - 1
        end = whole['first_row'] + whole['row_count'] - 1
    if not c1 <= pk_number <= c2:
        raise ValueError(f'表C扫描区域必须包含主键列 {pk_name}')
    skipped = effective_skip_columns(config, c1, c2)
    pk_offset = pk_number - c1
    spans = list(kept_spans(c1, c2, skipped))

    # 镜像位置（文件路径 + 工作表名）发生变化时：删除新位置上的旧镜像，强制重建一遍基线。
    mirror_path_value, signature = _mirror_signature(copier, config)
    if _previous_mirror_signature(copier) != signature and Path(mirror_path_value).exists():
        _reset_mirror_at(Path(mirror_path_value))

    mirror = _new_mirror(config, c1, c2, copier)
    activity = getattr(copier, 'activity', None)
    if activity is not None:
        activity.record('开始构建表A1镜像', path=str(mirror_path_value),
                        sheet=str(config.get('local', {}).get('sheet_name')),
                        columns=f'{c1}～{c2}', rows=f'{start}～{end}')
    try:
        mirror.ensure_table()
        if mirror.initialized:
            return dict(scanned=0, uploaded=0, pending=mirror.pending_count(), message='表A1完整镜像已存在，无需重建。')
        state = mirror.load_state()
    except Exception:
        mirror.close()
        raise

    seen = {}
    scanned = uploaded = blank = 0

    def source_batches():
        if whole is not None:
            # 整表已一次复制完成；按批切片仅用于进度与断点，不再重复复制。
            for bs, be in _row_batches(start, end, batch_rows, None):
                idx0 = bs - start
                rows = whole['rows'][idx0:idx0 + (be - bs + 1)]
                yield bs, be, rows, None
            return
        for bs, be in _row_batches(start, end, batch_rows, None):
            ref_c = f'{column_name(c1)}{bs}:{column_name(c2)}{be}'
            yield bs, be, _read_kept_columns(copier, page_c, ref_c, max_cells, skipped)['rows'], ref_c

    try:
        for batch_start, batch_end, rows, ref_c in source_batches():
            if _stop_requested(copier):
                raise SyncStopped()
            copier.update(message=(f'OverWatcher正在构建镜像：处理表C第{batch_start}～{batch_end}行（已上传{uploaded}行）…'
                                   if whole is not None else
                                   f'OverWatcher正在构建镜像：整体读取表C第{batch_start}～{batch_end}行（已上传{uploaded}行）…'),
                          scan_row=batch_start, scan_end=end, batch_end=batch_end)
            scanned += len(rows)
            upserts = []
            for offset, row in enumerate(rows):
                source_row = batch_start + offset
                key = pk_key(row[pk_offset])
                if key is None:
                    blank += 1
                    continue
                if key in seen:
                    raise ValueError(f'表C主键列 {pk_name} 存在重复值“{key}”（第{seen[key]}行与第{source_row}行），已停止本轮')
                seen[key] = source_row
                row = mirror_values(row, pk_offset)
                h = row_hash(row)
                prev = state.get(key)
                if prev is None or prev['row_hash'] != h:
                    upserts.append((key, row, h))
            if upserts:
                mirror.upsert(upserts)
                uploaded += len(upserts)
            if activity is not None:
                activity.record('构建镜像批次', rows=f'{batch_start}～{batch_end}',
                                upserts=len(upserts), scanned=scanned, uploaded=uploaded)
            if ref_c is not None:
                page_c.bring_to_front()
                copier._select(page_c, ref_c if not skipped else
                               f'{column_name(spans[0][0])}{batch_start}:{column_name(spans[0][1])}{batch_end}')
        mirror.finish_build()
        copier._a1_source_rows = dict(seen)
        _save_mirror_signature(copier, signature)
    finally:
        try:
            pending = mirror.pending_count()
        except Exception:
            pending = None
        mirror.close()
    result = dict(scanned=scanned, uploaded=uploaded, blank=blank, pending=pending,
                  batch_rows=batch_rows, skipped_columns=format_skip_columns(skipped))
    result['message'] = (f'OverWatcher第一阶段完成：已整体扫描表C {scanned}行，'
                         f'更新表A1镜像{uploaded}行，空主键{blank}行。'
                         + (f'镜像中尚有{pending}行待写回表A。' if pending else ''))
    if activity is not None:
        activity.record('表A1镜像构建完成', path=str(mirror_path_value),
                        scanned=scanned, uploaded=uploaded, blank=blank, pending=pending)
    copier.update(message=result['message'])
    return result


def run_sync(copier, config, only_rows=None, source_snapshot=None, reuse_target_index=None):
    """读一批、按主键比对并写完这一批，再继续；全表结束后由调度器等待。

    新流程：表C → 本地 表A1 镜像（变更检测） → 表A。
    - 每行与 表A1 镜像的 row_hash / pushed_hash 比对，区分“内容变化”（更新镜像）
      与“待写回”（进入同步候选）。
    - 候选行再与表A 当前内容比对，真正有差异才复制粘贴写回；成功后标记为已推送。

    only_rows：要同步的指定行号列表（用于“改动即同步”的即时路径）；为 None 时全表扫描。
    """
    if _stop_requested(copier):
        raise SyncStopped()
    pk_name, pk_number = parse_pk_column(config)
    batch_rows = sync_batch_rows(config)
    data_start = data_start_row(config)
    max_cells = config['source']['max_cells']
    # 每轮最多同步多少个变更行；0 = 不限（一轮扫完整表）。正数 N：写完 N 行就结束本轮。
    max_pastes = int(config.get('sync', {}).get('max_pastes', 0) or 0)
    page_c = _open_sheet(copier, config, 'c')
    whole = None
    if legacy_scan(config) or only_rows is not None or source_snapshot is not None:
        c1, start, c2, end = (source_snapshot['bounds'] if source_snapshot is not None
                              else _bounds(copier, page_c, config))
    else:
        # 方法2：整轮扫描（检索差异）也从下载的文件读取整表，与建镜像同源。
        whole = read_source_file(copier, page_c, config, '表C')
        c1, start = whole['first_col'], whole['first_row']
        c2 = whole['first_col'] + whole['width'] - 1
        end = whole['first_row'] + whole['row_count'] - 1
    if not c1 <= pk_number <= c2:
        raise ValueError(f'表C扫描区域必须包含主键列 {pk_name}')
    skipped = effective_skip_columns(config, c1, c2)
    # 表A 当前内容与主键→行号都来自表B（表A 整表下载快照）：默认（非旧检索方式）一律
    # “直接扫描全表”，不再在线分批读取表A索引/内容。旧检索方式（legacy_scan）仍按
    # scan_range + batch_rows 在线读取表A，保留为回退路径。
    table_b = None
    table_b_merges = set()
    if not legacy_scan(config):
        table_b = read_table_b_source(copier, config, c1, c2)
        table_b_merges = set(getattr(copier, '_table_b_merges', None) or ())
    # 表C确实存在待写回行时才打开表A和建立索引。
    page_a = None
    pk_offset, width = pk_number - c1, c2 - c1 + 1
    spans = list(kept_spans(c1, c2, skipped))

    def ref(first, last):
        return f'{column_name(c1)}{first}:{column_name(c2)}{last}'

    def target_map():
        if not legacy_scan(config):
            # 新方式：直接选中主键列（起始行 → 表A 最后有数据行）复制一次。
            return read_pk_column_index(copier, page_a, config, pk_number, pk_name, '表A')
        scan_range = config['upload'].get('scan_range') or None
        if not scan_range and skipped:
            scan_range = copier._used_range(page_a)
        if scan_range:
            left, right = scan_range.split(':')
            ac1, ar1 = address(left)
            ac2, ar2 = address(right)
            if not ac1 <= pk_number <= ac2:
                raise ValueError(f'表A扫描区域必须包含主键列 {pk_name}')
            # 索引只需要主键，避免其他隐藏列、格式和宽表复制限制影响索引。
            scan_range = f'{pk_name}{ar1}:{pk_name}{ar2}'
        grid = read_online_grid(copier, page_a, config, scan_range, '表A主键')
        first = grid['first_column']
        if not first <= pk_number < first + grid['width']:
            raise ValueError(f'表A扫描区域必须包含主键列 {pk_name}')
        index = build_index(grid, pk_name, pk_number, data_start,
                            config['upload'].get('scan_range') or '', label='表A')
        return {key: row_no for key, (_row, row_no) in index['rows'].items()}

    # 复用已建立的表A主键索引（表A行位置只被覆盖、不增删，因此索引稳定；
    # 若表A被人为改动导致错位，下方 moved 检测会重建）。
    cache_key = (config['upload'].get('share_url'), config['upload'].get('sheet_name'),
                 config['upload'].get('scan_range'), pk_name, data_start)
    if reuse_target_index is None:
        # 默认只让“改动即同步”的即时路径复用；筛选路径由调用方按“5 轮行号 + 1 轮主键
        # 定位”周期决定是否复用，从而与表C行号缓存保持同一节奏。
        reuse_target_index = only_rows is not None and source_snapshot is None
    a_map = None
    if reuse_target_index and getattr(copier, '_target_index_url', None) == cache_key:
        a_map = copier._target_index
    # 本地 表A1 镜像：建表并读取上一轮状态。
    mirror = _new_mirror(config, c1, c2, copier)
    try:
        mirror.ensure_table()
        state = mirror.load_state()
    except Exception:
        mirror.close()
        raise

    seen = {}
    scanned = blank = uploaded = 0
    copier._a1_target_keys = set()
    source_rows = {} if only_rows is None else dict(getattr(copier, '_a1_source_rows', {}))
    source_keys = {row: key for key, row in source_rows.items()}
    copier._a1_source_rows = source_rows
    # 表A 检查交给 WatchDog：OverWatcher 确认差异后，WatchDog 整体检索表A、按主键覆盖差异单元格。
    watchdog = WatchDog(copier, config, mirror, page_c, page_a, c1, c2, pk_name, pk_number,
                        pk_offset, width, spans, skipped, max_cells, max_pastes,
                        a_map, target_map, cache_key, table_b=table_b,
                        table_b_merges=table_b_merges)
    try:
        for batch_start, batch_end in _row_batches(start, end, batch_rows, only_rows):
            if _stop_requested(copier):
                raise SyncStopped()
            ref_c = ref(batch_start, batch_end)
            if whole is not None:
                # 整表已一次复制完成：按批切片只是便于进度、断点与 max_pastes，不再重复复制。
                idx0 = batch_start - start
                wanted = dict(rows=whole['rows'][idx0:idx0 + (batch_end - batch_start + 1)],
                              row_count=batch_end - batch_start + 1)
            elif source_snapshot is not None:
                wanted = dict(rows=[source_snapshot['rows'][no] for no in range(batch_start, batch_end+1)],
                              row_count=batch_end-batch_start+1)
            else:
                copier.update(message=f'正在读取表C第{batch_start}～{batch_end}行（每批{batch_rows}行，已更新{watchdog.changed}行）…',
                              scan_row=batch_start, scan_end=end, batch_end=batch_end)
                wanted = _read_kept_columns(copier, page_c, ref_c, max_cells, skipped)
            scanned += wanted['row_count']
            entries = []  # 需要写回表A 的行（内容相对“上次已推送”发生变化）
            upserts = []  # 内容相对镜像发生变化、需要更新 表A1 的行
            for offset, row in enumerate(wanted['rows']):
                source_row = batch_start + offset
                key = pk_key(row[pk_offset])
                old_key = source_keys.get(source_row)
                if old_key is not None and old_key != key:
                    source_rows.pop(old_key, None)
                if key is not None:
                    source_rows[key] = source_row
                    source_keys[source_row] = key
                if key is None:
                    blank += 1
                    continue
                if key in seen:
                    raise ValueError(f'表C主键列 {pk_name} 存在重复值“{key}”（第{seen[key]}行与第{source_row}行），已停止本轮')
                seen[key] = source_row
                row = mirror_values(row, pk_offset)
                h = row_hash(row)
                prev = state.get(key)
                content_changed = prev is None or prev['row_hash'] != h
                # 只按内容差异决定是否写回表A：A1 镜像（pushed_hash）与表C 当前内容（h）
                # 不同才推送；不再因为“格式待重试”而把内容已一致的单元格再次粘贴。
                needs_push = (prev is None or prev['pushed_hash'] != h)
                if content_changed:
                    upserts.append((key, row, h))
                if needs_push:
                    entries.append(dict(pk=key, source=source_row, row=row, hash=h))
            if upserts:
                mirror.upsert(upserts)
                uploaded += len(upserts)
                mirror.flush()

            if entries:
                copier._a1_target_keys.update(entry['pk'] for entry in entries)
                saved = mirror.read_rows([entry['pk'] for entry in entries])
                for entry in entries:
                    entry['row'] = saved[entry['pk']]
                if table_b is None:
                    # 旧检索方式：先打开表A并在线建立主键→行号索引。
                    if page_a is None:
                        page_a = _open_sheet(copier, config, 'a')
                        if a_map is None:
                            # 表A标签页被关掉/崩溃时（TargetClosedError）重新打开并重试一次，
                            # 避免整轮同步因浏览器偶发崩溃而中断。
                            for _attempt in range(2):
                                try:
                                    a_map = target_map()
                                    break
                                except Exception as exc:
                                    if type(exc).__name__ != 'TargetClosedError' or _attempt:
                                        raise
                                    copier.update_upload(message='表A页面已关闭或崩溃，正在重新打开表A并重试主键扫描…')
                                    page_a = _open_sheet(copier, config, 'a')
                            copier._target_index = a_map
                            copier._target_index_url = cache_key
                        watchdog.page_a = page_a
                        watchdog.a_map = a_map
                # OverWatcher 确认差异后，把变化行交给 WatchDog 整体检索表A并覆盖差异单元格。
                # 本地对比（table_b）时由 WatchDog 先本地找差异，确有差异才惰性打开表A粘贴。
                watchdog.process_entries(entries)
                if watchdog.stopped_early:
                    break
            copier.update(message=f'已完成表C第{batch_start}～{batch_end}行，累计扫描{scanned}行、更新表A1镜像{uploaded}行、更新表A{watchdog.changed}行；继续下一批…',
                          scan_row=batch_end, scan_end=end, batch_end=batch_end)
            activity = getattr(copier, 'activity', None)
            if activity is not None:
                activity.record('扫描表C批次', rows=f'{batch_start}～{batch_end}',
                                scanned=scanned, uploaded=uploaded, changed=watchdog.changed)
            # 只在“在线逐批读取表C”时才回到表C并选中下一批区域；整表文件读取（whole）
            # 与筛选快照（source_snapshot）都已本地完成，选中表C纯属多余，且会产生
            # 大量“选择单元格”的逐行扫描日志。
            if whole is None and source_snapshot is None:
                page_c.bring_to_front()
                copier._select(page_c, ref_c if not skipped else
                               f'{column_name(spans[0][0])}{batch_start}:{column_name(spans[0][1])}{batch_end}')
        page_c.bring_to_front()
    finally:
        try:
            pending = mirror.pending_count()
        except Exception:
            pending = None
        try:
            mirror.close()
        finally:
            watchdog.close()
    changed = watchdog.changed
    missing = watchdog.missing
    paste_batches = watchdog.paste_batches
    stopped_early = watchdog.stopped_early
    result = dict(baseline=False, scanned=scanned, modified=changed, added=0, deleted=0,
                  rows=changed, writes=changed, deletes=0, missing=missing, blank=blank,
                  uploaded=uploaded, pending=pending,
                  batch_rows=batch_rows, paste_batches=paste_batches, max_pastes=max_pastes,
                  stopped_early=stopped_early, targeted=only_rows is not None and source_snapshot is None,
                  skipped_columns=format_skip_columns(skipped))
    if stopped_early:
        tail = '本轮已按“每轮最多同步行数”设置结束，等待后重新从表C起始位置扫描。'
    elif only_rows is not None and source_snapshot is None:
        tail = '继续监听表C变化。'
    else:
        tail = '本轮扫描结束，等待后从表C起始位置重新扫描。'
    result['message'] = (f'{"即时同步" if only_rows is not None and source_snapshot is None else "本轮扫描"}完成：处理{scanned}行，'
                         f'更新表A1镜像{uploaded}行，更新表A {changed}行（{paste_batches}次粘贴）'
                         f'；表A无对应主键{missing}行，空主键{blank}行。'
                         + (f'表A1尚有{pending}行待写回表A。' if pending else '') + tail)
    if not changed and not missing:
        note = ('未扫描到表C内容变化，无需更新表A。' if not uploaded else
                '表C内容已更新表A1镜像，表A无需改动。')
        result['message'] = note + '未执行任何粘贴。' + result['message']
    copier.update(message=result['message'])
    return result


def _mode(values):
    """返回序列中出现次数最多的值；空序列返回 None。用于把页面提取的
    最常见行高/列宽当作表C的默认尺寸，避免把未显式声明的行/列重置错。"""
    from collections import Counter
    if not values:
        return None
    return Counter(values).most_common(1)[0][0]


def download_table_c(copier, config):
    """Download the original online workbook into the program's DATA folder."""
    from table_c_download import download_table_c as download
    return download(copier, config)


def download_table_b(copier, config):
    """下载表A（线上目标表）为 DATA/表B.xlsx，方法照搬表C。"""
    from table_c_download import download_table_b as download
    return download(copier, config)


def apply_table_c_format(copier, page_a, config):
    """把表C（DATA/表C.xlsx）的行高/列宽应用到线上表A。

    只应用行高/列宽（不会清空数据）；不自动应用合并单元格，因为合并会
    丢弃非左上角单元格数据，可能把整行清成空白。

    返回 None 表示未配置表C或无需处理；返回 dict 包含应用统计与提示信息。
    """
    from format_reference import load as load_reference
    from upload_layout import plan_layout
    from online_format import OnlineFormatter

    reference = load_reference(copier.root, config)
    if reference is None:
        return None
    source_range = config['source'].get('scan_range') or None
    scan_range = config['upload'].get('scan_range') or None
    if not source_range or not scan_range:
        return None
    destination = scan_range.split(':')[0]
    try:
        plan = plan_layout(reference, source_range, destination)
    except ValueError:
        return dict(error='表C格式范围与表C/表A扫描范围不匹配，本次未调整表A格式')
    if not plan.get('column_groups') and not plan.get('row_groups') and not plan.get('merges'):
        return None
    page_a.bring_to_front()
    formatter = OnlineFormatter(
        page_a, lambda ref: copier._select(page_a, ref),
        progress=lambda text: copier.update_upload(
            message='正在按表C设置表A格式：' + str(text).replace('表C', '表A')),
        diagnostic_dir=copier.root / 'error_reports')
    formatter.select_sheet(config['upload']['sheet_name'])
    formatter.preflight(plan)
    dimension_stats = formatter.apply_dimensions(plan)
    # 只应用行高/列宽；不自动应用合并单元格——合并会丢弃非左上角单元格数据，
    # 若按页面读取的合并坐标误作用到有数据的行，会把整行清成空白。
    return dict(sheet=reference.get('sheet'), merges=0,
                dimensions=dimension_stats,
                message=(f'已按表C设置表A格式：列宽{dimension_stats.get("column_groups", 0)}组、'
                         f'行高{dimension_stats.get("row_groups", 0)}组。'))
