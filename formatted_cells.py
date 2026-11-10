"""Table C key lookup and durable format retry state."""
import json
import os
from pathlib import Path

from sync_engine import pk_key, parse_pk_column, data_start_row


def file_version(path):
    stat = path.stat()
    return stat.st_mtime_ns, stat.st_size


class FormatPending:
    """Keep cells pending across a crash/value-success/format-failure boundary."""
    def __init__(self, root, config):
        self.path = Path(root) / 'sync_format_pending.json'
        self.identity = {name: {field: config.get(name, {}).get(field)
                              for field in ('share_url', 'sheet_name')}
                         for name in ('source', 'upload')}
        self.identity = dict(self.identity, pk=parse_pk_column(config)[0])
        self.cells = {}
        if self.path.exists():
            data = json.loads(self.path.read_text(encoding='utf-8'))
            if data.get('identity') == self.identity:
                self.cells = {key: set(cols) for key, cols in data['cells'].items()}

    def update(self, keys, left, right, pending):
        for key in keys:
            columns = self.cells.setdefault(key, set())
            if pending:
                columns.update(range(left, right + 1))
            else:
                columns.difference_update(range(left, right + 1))
            if not columns:
                self.cells.pop(key)
        temporary = self.path.with_suffix('.tmp')
        temporary.write_text(json.dumps(dict(identity=self.identity,
            cells={key: sorted(cols) for key, cols in self.cells.items()}), ensure_ascii=False), encoding='utf-8')
        os.replace(temporary, self.path)


class TableCIndex:
    def __init__(self, root, config):
        from openpyxl import load_workbook
        from format_reference import load
        source = config['source']
        self.path = (Path(root) / (source.get('format_file') or 'DATA/表C.xlsx')).resolve()
        self.sheet_name = source.get('format_sheet') or source['sheet_name']
        if not self.path.is_file():
            raise ValueError('必须先下载表C，才能按同主键同列复制格式；本次尚未粘贴表A')
        self.version = file_version(self.path)
        from zipfile import ZipFile
        with ZipFile(self.path) as archive:
            self.theme = archive.read('xl/theme/theme1.xml') if 'xl/theme/theme1.xml' in archive.namelist() else None
        self.rows = {}
        _, pk_col = parse_pk_column(config)
        book = load_workbook(self.path, read_only=True, data_only=True)
        try:
            if self.sheet_name not in book.sheetnames:
                raise ValueError('表C中找不到工作表“' + self.sheet_name + '”，请重新下载表C')
            sheet = book[self.sheet_name]
            for cells in sheet.iter_rows(min_row=data_start_row(config), min_col=pk_col, max_col=pk_col):
                cell = cells[0]
                value = cell.value
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    if float(value).is_integer():
                        value = str(int(value))
                        fmt = cell.number_format
                        if fmt and set(fmt) == {'0'}:
                            value = value.zfill(len(fmt))
                key = pk_key(value)
                if key is None:
                    continue
                if key in self.rows:
                    raise ValueError(f'表C主键“{key}”重复（第{self.rows[key]}、{cell.row}行），无法确定格式来源')
                self.rows[key] = cell.row
        finally:
            book.close()
        reference = load(root, {'source': {'format_file': str(self.path), 'format_sheet': self.sheet_name}})
        self.merges = reference['merges']
        self.check()

    def check(self):
        if file_version(self.path) != self.version:
            raise ValueError('表C在同步期间被修改，请重新同步；本次尚未继续粘贴')

    def resolve(self, keys, left, right):
        self.check()
        rows = []
        for key in keys:
            if key not in self.rows:
                raise ValueError(f'表C找不到主键“{key}”，请下载最新表C；不能套用其他行格式')
            row = self.rows[key]
            if any(top <= row <= bottom and a <= right and b >= left
                   for a, top, b, bottom in self.merges):
                raise ValueError(f'表C主键“{key}”对应列涉及合并单元格，无法仅改变差异单元格，请先取消该处合并')
            rows.append(row)
        return rows
