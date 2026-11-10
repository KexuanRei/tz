"""Read live rules each round and synchronize an independent filtered snapshot."""
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path
import json
import re

from copy_range import column_name
from local_mirror import (LocalMirror, mirror_path, mirror_values, STATE_SHEET,
                          mirror_identity, validate_mirror_identity)
from sync_engine import parse_pk_column, pk_key, sync_batch_rows, legacy_scan
from sync_columns import effective_skip_columns
from filter_headers import read_header_row, header_name


def enabled(config):
    return config.get('filter', {}).get('enabled', False)


def effective_config(root, config):
    cfg = deepcopy(config)
    options = cfg.get('filter', {})
    path = (Path(root) / options.get('local_file', 'DATA/筛选镜像.xlsx')).resolve()
    protected = [mirror_path(root, config),
                 (Path(root) / (config['source'].get('format_file') or 'DATA/表C.xlsx')).resolve(),
                 (Path(root) / 'DATA/表C.xlsx').resolve()]
    if path in protected:
        raise ValueError('筛选镜像必须使用独立文件，不能与整表镜像或表C相同')
    if path.suffix.lower() != '.xlsx':
        raise ValueError('筛选镜像路径必须指向.xlsx文件')
    cfg['local'] = dict(local_file=str(path), sheet_name=options.get('mirror_sheet', '筛选镜像'))
    # 筛选镜像沿用整表同步的读取方式：默认（非旧检索方式）表C 从下载的整表文件读取、
    # 表A 从表B（表A 整表快照）本地对比，即“直接扫描全表”；旧检索方式仍在线分批读取。
    return cfg


def parse_rules(rows):
    if not rows:
        raise ValueError('系统配置为空，已停止本轮，不会回退为整表同步')
    header = [str(v or '').strip() for v in rows[0]]
    if any(header.count(name) != 1 for name in ('关键列', '条件')):
        raise ValueError('系统配置第1行必须有且仅有一个“关键列”和“条件”表头')
    key_col, value_col = header.index('关键列'), header.index('条件')
    result = []
    for row_no, row in enumerate(rows[1:], 2):
        row = list(row) + [''] * len(header)
        key, value = str(row[key_col] or '').strip(), str(row[value_col] or '').strip()
        if not key and not value:
            continue
        if not key or not value:
            raise ValueError(f'系统配置第{row_no}行的关键列和条件必须同时填写')
        result.append((key, value))
    if not result:
        raise ValueError('系统配置没有有效条件，已停止本轮，不会回退为整表同步')
    return result


def rules_signature(rules):
    """Deterministic signature of the live filter rules, persisted in the mirror.

    Changing 关键列 or 条件 must invalidate the established mirror so the next
    round performs a full rescan instead of only re-reading remembered rows.
    """
    return json.dumps(sorted(rules), ensure_ascii=False)


def matches(value, condition):
    """Month rules match explicit dates/month labels; all other rules are exact."""
    month = re.fullmatch(r'(0?[1-9]|1[0-2])月', condition)
    text = str(value if value is not None else '').strip()
    if month:
        target = int(month[1])
        if isinstance(value, (datetime, date)):
            return value.month == target
        # Explicit year/month/day, month/day, or Chinese month/day. Never substring-match 14月.
        match = re.fullmatch(r'(?:(\d{4})[-/.年])?(\d{1,2})(?:月(?:(\d{1,2})日?)?|[-/.](\d{1,2}))(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?', text)
        if not match or int(match[2]) != target:
            return False
        try:
            date(int(match[1] or 2000), int(match[2]), int(match[3] or match[4] or 1))
        except ValueError:
            return False
        return True
    return text == condition


def bind_rules(rules, header, first_col):
    result = []
    names = [header_name(value) for value in header]
    for name, condition in rules:
        normalized = header_name(name)
        if not normalized or names.count(normalized) != 1:
            raise ValueError(f'表C表头中的“{name}”不存在或重复，无法确定筛选列')
        result.append((first_col + names.index(normalized), condition))
    return result


def _header_row_local(root, config, c1, c2, header_row, skipped):
    """从下载的表C文件读取表头行；合并单元格取左上角值（等价于在线逐列识别）。

    合并定义用 XML 轻量读取、表头及合并左上角只读少数几行，避免整表载入内存。
    """
    from openpyxl import load_workbook
    from source_file import _data_path, _sheet_name
    skipped = set(skipped)
    path = _data_path(root, config)
    if not Path(path).is_file():
        raise ValueError('表C数据文件不存在：' + str(path) + '；请先下载表C')
    merges = ()
    try:
        from format_reference import load as load_layout
        merges = (load_layout(root, config) or {}).get('merges') or ()
    except Exception:
        merges = ()
    relevant = [(mc1, mr1, mc2, mr2) for mc1, mr1, mc2, mr2 in merges
                if mr1 <= header_row <= mr2]
    min_row = min([header_row] + [mr1 for _c1, mr1, _c2, _r2 in relevant])
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book[_sheet_name(config, book)]
        row_map = {}
        for idx, row in enumerate(sheet.iter_rows(min_row=min_row, max_row=header_row,
                                                  min_col=c1, max_col=c2, values_only=True),
                                  start=min_row):
            row_map[idx] = list(row)
    finally:
        book.close()
    header = row_map.get(header_row, [''] * (c2 - c1 + 1))
    names = []
    for i, col in enumerate(range(c1, c2 + 1)):
        if col in skipped:
            names.append('')
            continue
        value = header[i] if i < len(header) else None
        for mc1, mr1, mc2, mr2 in relevant:
            if mc1 <= col <= mc2:
                top = row_map.get(mr1, [])
                offset = mc1 - c1
                value = top[offset] if 0 <= offset < len(top) else None
                break
        names.append(value)
    return names


def _key_columns(copier, page, config, rules, c1, c2, header_row, skipped):
    """Map each rule's normalized 关键列 to its absolute column number.

    系统配置的关键列没有变化时直接复用上次映射，不再重读表头；重读时也只读到
    所有关键列都找到为止（合并表头的逐列识别可提前停止）。
    """
    key_names = tuple(sorted({header_name(name) for name, _ in rules if header_name(name)}))
    cache_key = (config['source'].get('share_url'), config['source'].get('sheet_name'),
                 key_names, c1, c2, header_row, tuple(sorted(skipped)))
    cache = getattr(copier, '_filtered_key_columns_cache', None)
    if cache is not None and cache[0] == cache_key:
        return cache[1]
    if not legacy_scan(config):
        header = _header_row_local(copier.root, config, c1, c2, header_row, skipped)
    else:
        header = read_header_row(copier, page, c1, c2, header_row, config['source']['max_cells'],
                                 skipped, stop_names=set(key_names) or None)
    names = [header_name(value) for value in header]
    columns = {}
    for name in key_names:
        if names.count(name) != 1:
            raise ValueError(f'表C表头中的“{name}”不存在或重复，无法确定筛选列')
        columns[name] = c1 + names.index(name)
    copier._filtered_key_columns_cache = (cache_key, columns)
    return columns


def _mirror_state(root, config):
    """Return (initialized, keys, rules) of the established filtered mirror.

    keys is a set of normalized primary keys when the mirror is initialized;
    otherwise keys is None. rules is the stored rules signature, or None when
    the mirror predates rules tracking.
    """
    from openpyxl import load_workbook
    path = mirror_path(root, config)
    if not path.exists():
        return False, None, None
    book = load_workbook(path, read_only=True)
    try:
        if STATE_SHEET not in book:
            return False, None, None
        meta = json.loads(book[STATE_SHEET].cell(1, 1).value or '{}')
        saved = meta.get('identity') or {}
        scan_range = config['source'].get('scan_range')
        if scan_range:
            from update_excel import address
            first, last = scan_range.split(':')
            first_col, _ = address(first)
            last_col, _ = address(last)
        else:
            first_col, last_col = saved.get('first_col'), saved.get('last_col')
        validate_mirror_identity(saved, mirror_identity(config, first_col, last_col))
        stored_rules = meta.get('rules')
        if not meta.get('initialized'):
            return False, None, stored_rules
        keys = set()
        for values in book[STATE_SHEET].iter_rows(min_row=3, values_only=True):
            if values and values[0] is not None:
                key = pk_key(values[0])
                if key is not None:
                    keys.add(key)
        return True, keys, stored_rules
    finally:
        book.close()


def _read_rules_local(root, config):
    """从下载的表C文件（DATA/表C.xlsx）读取「系统配置」工作表的筛选规则。"""
    from openpyxl import load_workbook
    from source_file import _data_path
    options = config['filter']
    sheet_name = options.get('sheet_name', '系统配置')
    path = _data_path(root, config)
    if not Path(path).is_file():
        raise ValueError('表C数据文件不存在：' + str(path) + '；请先下载表C')
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in book.sheetnames:
            raise ValueError(f'下载的表C中没有“{sheet_name}”工作表，无法读取系统配置')
        sheet = book[sheet_name]
        rows = [list(row) for row in sheet.iter_rows(values_only=True)]
    finally:
        book.close()
    # openpyxl 的 max_row 可能因残留格式被撑大，去掉尾部整行为空的行。
    while rows and all(v is None or str(v).strip() == '' for v in rows[-1]):
        rows.pop()
    return parse_rules(rows)


def _read_rules(copier, config):
    """Read the live 关键列/条件 rules from the 系统配置 sheet.

    默认（非旧检索方式）从下载的表C文件本地读取；旧检索方式仍在线读取。
    """
    if not legacy_scan(config):
        return _read_rules_local(copier.root, config)
    from sync_runner import (_open_sheet, _row_batches, _read_block,
                             _stop_requested, SyncStopped)
    options = config['filter']
    rules_config = deepcopy(config)
    rules_config['source']['sheet_name'] = options.get('sheet_name', '系统配置')
    page = _open_sheet(copier, rules_config, 'c')
    ref = copier._used_range(page)
    from update_excel import address
    c1, r1 = address(ref.split(':')[0])
    c2, r2 = address(ref.split(':')[1])
    if r1 != 1:
        raise ValueError('系统配置的“关键列”“条件”表头必须位于第1行')
    rules_rows = []
    for lo, hi in _row_batches(1, r2, sync_batch_rows(config)):
        if _stop_requested(copier):
            raise SyncStopped()
        rules_rows.extend(_read_block(copier, page, f'{column_name(c1)}{lo}:{column_name(c2)}{hi}',
                                      config['source']['max_cells'])['rows'])
    return parse_rules(rules_rows)


def read_snapshot(copier, config, rules=None):
    from sync_runner import (_open_sheet, _bounds, _row_batches, _read_kept_columns,
                             _stop_requested, SyncStopped, read_source_file, legacy_scan)
    if rules is None:
        rules = _read_rules(copier, config)
    page = _open_sheet(copier, config, 'c')
    whole = None
    if not legacy_scan(config):
        # 直接扫描全表：表C 从下载的整表文件读取，不再在线分批复制。
        whole = read_source_file(copier, page, config, '表C')
        c1, start = whole['first_col'], whole['first_row']
        c2 = whole['first_col'] + whole['width'] - 1
        end = whole['first_row'] + whole['row_count'] - 1
    else:
        c1, start, c2, end = _bounds(copier, page, config)
    _, pk = parse_pk_column(config)
    if not c1 <= pk <= c2:
        raise ValueError('表C扫描区域必须包含主键列')
    options = config['filter']
    header_row = options.get('header_row', 1)
    skipped = effective_skip_columns(config, c1, c2)
    columns = _key_columns(copier, page, config, rules, c1, c2, header_row, skipped)
    bound = [(columns[header_name(name)], condition) for name, condition in rules]
    selected, seen = {}, {}
    scanned = blank = 0
    for lo, hi in _row_batches(start, end, sync_batch_rows(config), None):
        if _stop_requested(copier):
            raise SyncStopped()
        copier.update(message=f'正在按系统配置筛选表C第{lo}～{hi}行，已命中{len(selected)}行…')
        if whole is not None:
            # 整表已一次读取完成：按批切片仅用于进度与断点，不再重复复制。
            idx0 = lo - start
            rows = whole['rows'][idx0:idx0 + (hi - lo + 1)]
        else:
            # Skip sync-excluded columns exactly like the normal mirror: they are never
            # read (a hidden exclusion would otherwise return zero cells), and stay blank.
            block = _read_kept_columns(copier, page, f'{column_name(c1)}{lo}:{column_name(c2)}{hi}',
                                       config['source']['max_cells'], skipped)
            rows = block['rows']
        for offset, row in enumerate(rows):
            scanned += 1
            key = pk_key(row[pk-c1])
            if key is not None:
                if key in seen:
                    raise ValueError(f'表C主键“{key}”重复，已停止筛选，本轮尚未写入表A')
                seen[key] = lo + offset
            if not all(matches(row[col-c1], condition) for col, condition in bound):
                continue
            if key is None:
                blank += 1
                continue
            values = [value if c1+i not in skipped else '' for i, value in enumerate(row)]
            selected[lo+offset] = mirror_values(values, pk-c1)
    return dict(rows=selected, bounds=(c1, start, c2, end), rules=rules, scanned=scanned, blank=blank)


def refresh_mirror(copier, config, snapshot):
    """Commit membership atomically, preserving successful pushes across refreshes."""
    from openpyxl import load_workbook
    import json
    path = mirror_path(copier.root, config)
    if path.exists():
        book = load_workbook(path, read_only=True)
        try:
            meta = json.loads(book[STATE_SHEET].cell(1, 1).value) if STATE_SHEET in book else {}
            if not meta.get('identity', {}).get('filtered'):
                raise ValueError('筛选镜像路径已有其他文件，请选择新的独立.xlsx路径')
        finally:
            book.close()
    c1, _, c2, _ = snapshot['bounds']
    mirror = LocalMirror(copier.root, config, c1, c2, getattr(copier, 'activity', None))
    try:
        baseline = not mirror.initialized
        mirror.replace_members(list(snapshot['rows'].values()), baseline=baseline,
                               rules=rules_signature(snapshot['rules']))
    finally:
        mirror.close()
    copier._a1_mirror = None
    copier._a1_source_rows = {pk_key(row[mirror.pk_offset]): no for no, row in snapshot['rows'].items()}
    activity = getattr(copier, 'activity', None)
    if activity is not None:
        activity.record('筛选镜像重建', path=str(mirror.path),
                        matched=len(snapshot['rows']), baseline=baseline)
    return baseline


def run_filtered(copier, config):
    from sync_runner import run_sync, _stop_requested, SyncStopped
    cfg = effective_config(copier.root, config)
    # 校验镜像身份（不兼容的源/范围会在扫描和写回前报错）。
    _mirror_state(copier.root, cfg)
    scan_key = json.dumps(dict(source=cfg['source'], local=cfg['local'],
        filter=cfg['filter'], sync={key: cfg.get('sync', {}).get(key) for key in
        ('primary_key_column', 'data_start_row', 'skip_columns')}), sort_keys=True)
    if getattr(copier, '_filtered_scan_key', None) != scan_key:
        copier._filtered_key_columns_cache = None
        copier._filtered_scan_key = scan_key
    # 非旧检索方式：先下载表C整表文件，之后规则/表头/数据都从本地文件读取，
    # 只有最后写回线上表A是在线操作（与整表同步一致）。
    if not legacy_scan(cfg):
        from sync_runner import read_source_file, _open_sheet
        read_source_file(copier, _open_sheet(copier, cfg, 'c'), cfg, '表C')
    rules = _read_rules(copier, cfg)
    # 每一轮都直接扫描全表：从本地表C整表按「关键列/条件」命中并重建筛选镜像，
    # 不做“按已记住行号/按主键定位”的增量扫描——那些增量路径虽然省了扫描量，
    # 但把本地全表扫描退化成逐行跳转，反而更慢，也永远发现不了新命中的行。
    snapshot = read_snapshot(copier, cfg, rules=rules)
    if _stop_requested(copier):
        raise SyncStopped()
    baseline = refresh_mirror(copier, cfg, snapshot)
    result = run_sync(copier, cfg, only_rows=list(snapshot['rows']), source_snapshot=snapshot)
    if cfg.get('sync', {}).get('insert_enabled') and result.get('missing'):
        from cell_runner import run_insert
        from edge_copy import _merge_insert_result
        result = _merge_insert_result(result, run_insert(copier, cfg, only_rows=list(snapshot['rows'])))
    result.update(filtered=True, filter_matched=len(snapshot['rows']), filter_scanned=snapshot['scanned'],
                  filter_baseline=baseline, filter_mirror=cfg['local']['local_file'])
    note = '首次建立筛选基线；后续变化开始同步。' if baseline else ''
    result['message'] = (f'系统配置筛选（整表扫描）：扫描{snapshot["scanned"]}行，命中{len(snapshot["rows"])}行；'
                         f'独立镜像：{cfg["local"]["local_file"]}。' + note + result['message'])
    copier.update(message=result['message'])
    return result
