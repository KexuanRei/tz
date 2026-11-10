"""Persistent .xlsx mirror: mutate only rows addressed by primary key."""
import json
import os
from pathlib import Path
import tempfile

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from copy_range import column_name
from sync_engine import parse_pk_column, pk_key, row_hash
from sync_columns import effective_skip_columns
from upload_verify import comparison_result

STATE_SHEET = '_sync_state'
FORMAT_VERSION = 1


def mirror_identity(config, first_col, last_col):
    _, primary_key = parse_pk_column(config)
    identity = dict(version=FORMAT_VERSION, first_col=first_col, last_col=last_col,
        primary_key=primary_key,
        sheet=str(config.get('local', {}).get('sheet_name') or '表A1'),
        source_url=config['source'].get('share_url'), source_sheet=config['source'].get('sheet_name'),
        scan_range=config['source'].get('scan_range') or '',
        data_start=config.get('sync', {}).get('data_start_row', 2),
        skipped=list(effective_skip_columns(config, first_col, last_col)))
    if config.get('filter', {}).get('enabled'):
        identity['filtered'] = True
    return identity


def validate_mirror_identity(saved, current):
    """Allow only row-window edits to explicit ranges with unchanged columns."""
    if not isinstance(saved, dict):
        raise ValueError('表A1镜像元数据损坏，请恢复备份')
    if saved == current:
        return False
    changed = sorted(key for key in set(saved or {}) | set(current)
                     if (saved or {}).get(key) != current.get(key))
    if changed == ['scan_range']:
        from update_excel import address
        def columns(identity):
            start, end = identity['scan_range'].split(':')
            left, top = address(start)
            right, bottom = address(end)
            if not (1 <= top <= bottom <= 1048576 and
                    (left, right) == (identity['first_col'], identity['last_col'])):
                raise ValueError('invalid scan range')
            return left, right
        try:
            if columns(saved) == columns(current):
                return True
        except (KeyError, TypeError, AttributeError, ValueError):
            pass
    names = {'source_url':'表C链接', 'source_sheet':'表C工作表', 'primary_key':'主键列',
             'first_col':'起始列', 'last_col':'结束列', 'scan_range':'扫描范围',
             'data_start':'数据起始行', 'skipped':'跳过列', 'sheet':'镜像工作表',
             'filtered':'筛选模式', 'version':'镜像版本'}
    detail = '、'.join(names.get(key, key) for key in changed)
    exc = ValueError(f'表A1镜像与当前配置不一致（{detail}），请恢复配置或选择新的.xlsx路径')
    exc.diagnostic_context = {'changed_fields': changed,
        'saved_scan_range': (saved or {}).get('scan_range'),
        'current_scan_range': current.get('scan_range')}
    raise exc


def mirror_path(root, config):
    value = str(config.get('local', {}).get('local_file') or '表A1.xlsx').strip()
    path = (Path(root) / value).resolve()
    if path.suffix.lower() != '.xlsx':
        raise ValueError('表A1路径必须指向.xlsx文件')
    return path


def mirror_values(row, pk_offset):
    # Keep identifiers literal, including leading zeroes. Store formula results,
    # never formulas referring to the source workbook's coordinates.
    return [str(value if i == pk_offset else comparison_result(value)) if value is not None else ''
            for i, value in enumerate(row)]


def _stamp(path):
    if not path.exists():
        return None
    s = path.stat()
    return s.st_mtime_ns, s.st_size


class LocalMirror:
    def __init__(self, root, config, first_col, last_col, activity=None):
        self.path = mirror_path(root, config)
        self.activity = activity
        self.first_col, self.last_col = first_col, last_col
        self.width = last_col - first_col + 1
        self.sheet_name = str(config.get('local', {}).get('sheet_name') or '表A1')
        if self.sheet_name == STATE_SHEET:
            raise ValueError('表A1工作表名称不能使用保留名称 _sync_state')
        _, self.pk_number = parse_pk_column(config)
        self.pk_offset = self.pk_number - first_col
        if not 0 <= self.pk_offset < self.width:
            raise ValueError('表A1镜像列范围必须包含主键列')
        self.identity = mirror_identity(config, first_col, last_col)
        self.state = {}
        self.dirty = False
        self.initialized = False
        self.stamp = _stamp(self.path)
        if self.stamp is None:
            self.book = Workbook()
            self._prepare_fresh_sheet()
        else:
            self.book = load_workbook(self.path)
            if STATE_SHEET in self.book and self.sheet_name in self.book:
                try:
                    self._load_existing_mirror()
                except Exception:
                    self.book.close()
                    raise
            else:
                # 用户自己指定的文件，尚未被本程序初始化：就地“收养”为新镜像。
                # 只重建目标数据工作表与 _sync_state，保留文件里其它工作表。
                self._adopt_user_file()
        self.next_row = max([v['row'] for v in self.state.values()] + [1]) + 1

    def _prepare_fresh_sheet(self):
        self.sheet = self.book.active
        self.sheet.title = self.sheet_name
        self._write_header_and_meta()

    def _adopt_user_file(self):
        if STATE_SHEET in self.book:
            self.book.remove(self.book[STATE_SHEET])
        if self.sheet_name in self.book:
            self.book.remove(self.book[self.sheet_name])
        self.sheet = self.book.create_sheet(self.sheet_name, 0)
        self._write_header_and_meta()

    def _write_header_and_meta(self):
        for col in range(self.first_col, self.last_col + 1):
            self.sheet.cell(1, col, column_name(col))
        self.sheet.freeze_panes = 'A2'
        self.meta = self.book.create_sheet(STATE_SHEET)
        self.meta.sheet_state = 'veryHidden'
        self.meta.append([json.dumps(dict(identity=self.identity, initialized=False), ensure_ascii=False)])
        self.meta.append(['primary_key', 'local_row', 'row_hash', 'pushed_hash'])

    def _load_existing_mirror(self):
        self.sheet, self.meta = self.book[self.sheet_name], self.book[STATE_SHEET]
        metadata = json.loads(self.meta.cell(1, 1).value or '{}')
        range_changed = validate_mirror_identity(metadata.get('identity'), self.identity)
        self.initialized = bool(metadata.get('initialized'))
        used_rows = set()
        for meta_row, values in enumerate(self.meta.iter_rows(min_row=3, values_only=True), 3):
            key, local_row, h, pushed = values[:4]
            if not key or key in self.state or not isinstance(local_row, int) or local_row < 2 or local_row in used_rows:
                raise ValueError('表A1镜像索引损坏或主键重复，请恢复备份')
            used_rows.add(local_row)
            row = self._values(local_row)
            if pk_key(row[self.pk_offset]) != key or row_hash(row) != h:
                raise ValueError('表A1镜像内容被外部修改，请恢复镜像或选择新路径重建')
            self.state[key] = dict(row=local_row, meta_row=meta_row, row_hash=h, pushed_hash=pushed)
        if range_changed:
            # Update metadata only after validating all stored values/hashes.
            # Keep initialization, local row positions and pending pushes intact.
            metadata['identity'] = self.identity
            self.meta.cell(1, 1, json.dumps(metadata, ensure_ascii=False))
            self.dirty = True

    def unchanged_on_disk(self):
        return _stamp(self.path) == self.stamp

    def ensure_table(self):
        if not self.unchanged_on_disk():
            raise ValueError('表A1文件在同步期间被外部修改，已停止以避免覆盖')

    def _values(self, row_no):
        values = []
        for col in range(self.first_col, self.last_col + 1):
            cell = self.sheet.cell(row_no, col)
            if cell.data_type == 'f':
                raise ValueError('表A1镜像应保存计算结果，不能包含手工公式')
            values.append('' if cell.value is None else str(cell.value))
        return values

    def load_state(self):
        return {key: dict(row_hash=v['row_hash'], pushed_hash=v['pushed_hash']) for key, v in self.state.items()}

    def read_rows(self, keys):
        self.ensure_table()
        result = {}
        for key in keys:
            if key not in self.state:
                raise ValueError(f'表A1没有主键“{key}”')
            state = self.state[key]
            row = self._values(state['row'])
            if pk_key(row[self.pk_offset]) != key or row_hash(row) != state['row_hash']:
                raise ValueError(f'表A1主键“{key}”的内容与索引不一致，已停止')
            result[key] = row
        return result

    def key_rows(self):
        return sorted((v['row'], key) for key, v in self.state.items())

    def pending_count(self):
        return sum(v['row_hash'] != v['pushed_hash'] for v in self.state.values())

    def upsert(self, rows):
        self.ensure_table()
        prepared = []
        for key, values, h in rows:
            values = mirror_values(values, self.pk_offset)
            if len(values) != self.width or pk_key(values[self.pk_offset]) != key or row_hash(values) != h:
                raise ValueError('写入表A1的主键、列宽或内容哈希不一致')
            if any(len(v) > 32767 or ILLEGAL_CHARACTERS_RE.search(v) for v in values):
                raise ValueError('表A1单元格超过Excel文本长度限制或含不支持字符')
            prepared.append((key, values, h))
        for key, values, h in prepared:
            previous = self.state.get(key)
            if previous and previous['row_hash'] == h:
                continue
            if previous is None:
                previous = dict(row=self.next_row, meta_row=len(self.state)+3, row_hash=None, pushed_hash=None)
                self.next_row += 1
                self.state[key] = previous
            for offset, value in enumerate(values):
                cell = self.sheet.cell(previous['row'], self.first_col+offset)
                if (cell.value or '') != value:
                    cell.value = value
                    cell.data_type = 's'
            previous['row_hash'] = h
            for col, value in enumerate((key, previous['row'], h, previous['pushed_hash']), 1):
                cell = self.meta.cell(previous['meta_row'], col)
                cell.value = value
                if isinstance(value, str):
                    cell.data_type = 's'
            self.dirty = True

    def mark_pushed(self, entries):
        for key, h in entries:
            state = self.state.get(key)
            if state is None or state['row_hash'] != h:
                raise ValueError('表A1内容已变化，不能标记旧版本为已同步')
            if state['pushed_hash'] != h:
                state['pushed_hash'] = h
                self.meta.cell(state['meta_row'], 4, h)
                self.dirty = True
        self.flush()

    def finish_build(self):
        # First construction is a baseline, not a request to overwrite table A.
        self.initialized = True
        self.meta.cell(1, 1, json.dumps(dict(identity=self.identity, initialized=True), ensure_ascii=False))
        self.dirty = True
        self.mark_pushed([(key, v['row_hash']) for key, v in self.state.items()])

    def replace_members(self, rows, baseline=False, rules=None):
        """Replace filtered membership, retaining push state only for surviving keys."""
        self.ensure_table()
        previous = self.load_state()
        self.sheet.delete_rows(2, max(0, self.sheet.max_row - 1))
        self.meta.delete_rows(3, max(0, self.meta.max_row - 2))
        self.state, self.next_row = {}, 2
        self.dirty = True
        prepared = [(pk_key(row[self.pk_offset]), row, row_hash(row)) for row in rows]
        self.upsert(prepared)
        for key, state in self.state.items():
            state['pushed_hash'] = state['row_hash'] if baseline else previous.get(key, {}).get('pushed_hash')
            self.meta.cell(state['meta_row'], 4).value = state['pushed_hash']
        self.initialized = True
        meta = dict(identity=self.identity, initialized=True)
        if rules is not None:
            meta['rules'] = rules
        self.meta.cell(1, 1, json.dumps(meta, ensure_ascii=False))
        self.flush()

    def flush(self):
        if not self.dirty:
            return
        self.ensure_table()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        handle = tempfile.NamedTemporaryFile(dir=self.path.parent, suffix='.xlsx', delete=False)
        temp = Path(handle.name)
        handle.close()
        try:
            self.book.save(temp)
            self.ensure_table()
            os.replace(temp, self.path)
        except PermissionError as exc:
            raise ValueError('表A1无法保存，请关闭正在打开此文件的Excel/WPS后重试') from exc
        finally:
            temp.unlink(missing_ok=True)
        self.stamp = _stamp(self.path)
        self.dirty = False
        if self.activity is not None:
            try:
                size = self.path.stat().st_size
            except OSError:
                size = None
            self.activity.record('保存表A1镜像', path=str(self.path), sheet=self.sheet_name,
                                 rows=len(self.state), bytes=size)

    def close(self):
        # Saves are explicit, before publishing values and after verification.
        # Cache only a committed workbook. An interrupted build is discarded.
        self.book.close()
