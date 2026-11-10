"""表C→表A1→表A 同步的纯逻辑：主键索引、变化检测、表A1增量写入。

本模块不接触浏览器、剪贴板或 WPS/Excel COM，只处理已经读到的网格数据和
本地文件，便于单元测试与复用。变化检测以主键（默认第4列=列D）匹配数据行：

- 新增：本轮主键存在、上一轮不存在；
- 修改：主键相同但行内容有差异（忽略纯空白差异）；
- 删除：上一轮主键存在、本轮不存在。

上一轮快照只保存每行的哈希（不含单元格内容），避免大表占用过多磁盘。
"""
import datetime as dt
import hashlib
import json
from pathlib import Path

from copy_range import column_name
from update_excel import column, address
from upload_verify import normalize_text

SNAPSHOT_FILE = 'sync_snapshot.json'
DELTA_FILE = 'sync_delta.json'
STATUS_LABELS = {'add': '新增', 'modify': '修改', 'delete': '删除'}


def parse_pk_column(config):
    """返回 (列名, 列号) 的元组；默认第4列=列D。"""
    name = str(config.get('sync', {}).get('primary_key_column') or 'D').strip().upper()
    number = column(name)
    return name, number


def data_start_row(config):
    """数据起始行（1基，第1行为表头）。默认2。"""
    value = config.get('sync', {}).get('data_start_row', 2)
    if not isinstance(value, int) or value < 1:
        raise ValueError('数据起始行必须是大于等于1的整数（第1行为表头时填2）')
    return value


def start_cell(config):
    """整表复制的起始单元格，返回 (列名, 列号, 行号)。默认 A5。"""
    value = str(config.get('sync', {}).get('start_cell') or 'A5').strip().upper().replace('$', '')
    if ':' in value:
        raise ValueError('起始单元格必须是单个单元格地址，例如 A5')
    try:
        number, row = address(value)
    except (ValueError, TypeError):
        raise ValueError('起始单元格须是单元格地址，例如 A5（列字母 + 行号）')
    if row < 1:
        raise ValueError('起始单元格行号必须大于等于1')
    return column_name(number), number, row


def legacy_scan(config):
    """是否沿用旧的分批范围读取（scan_range + batch_rows）。默认关闭（整表复制）。"""
    return bool(config.get('sync', {}).get('legacy_scan', False))


MAX_BATCH_ROWS = 5000  # 一次读取/同步的最大行数上限


def sync_batch_rows(config):
    """一次读取、比较和处理的行数，不限制整轮扫描范围。"""
    value = config.get('sync', {}).get('batch_rows', 100)
    if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= MAX_BATCH_ROWS:
        raise ValueError(f'每批同步行数须为1到{MAX_BATCH_ROWS}的整数')
    return value


def pk_key(value):
    """把主键值归一化为字符串键；空值返回 None（该行不参与匹配）。"""
    if value is None:
        return None
    text = normalize_text(value)
    return text or None


def grid_size(grid):
    """网格列宽：header 与所有行里最宽者。"""
    width = grid.get('width') or 0
    if not width:
        width = max([len(grid.get('header') or [])] + [len(r) for r in (grid.get('rows') or [])] or [0])
    return width


def build_index(grid, pk_name, pk_number, data_start, scan_range, label='表C'):
    """从扫描网格构建本轮主键索引（含完整行值，仅内存使用）。"""
    first_col = grid.get('first_column') or 1
    first_row = grid.get('first_row') or 1
    pk_index = pk_number - first_col
    header = list(grid.get('header') or [])
    rows = {}
    for offset, row in enumerate(grid.get('rows') or []):
        sheet_row = first_row + offset
        if sheet_row < data_start:
            continue
        key = pk_key(row[pk_index] if 0 <= pk_index < len(row) else '')
        if key is None:
            continue
        if key in rows:
            raise ValueError(
                f'{label}主键列 {pk_name} 存在重复值“{key}”（第{rows[key][1]}行与第{sheet_row}行），'
                f'无法唯一匹配。请更换主键列，或修正{label}中的重复值。')
        rows[key] = (row, sheet_row)
    return dict(primary_key_column=pk_name, data_start_row=data_start, range=scan_range,
                width=grid_size(grid), first_col=first_col, first_row=first_row,
                header=header, rows=rows)


def row_hash(row):
    normalized = '\x1f'.join(normalize_text(v) for v in row)
    return hashlib.sha256(normalized.encode('utf-8')).hexdigest()


def index_hashes(index):
    return {key: row_hash(row) for key, (row, _row_no) in index['rows'].items()}


def diff_index(prev_hashes, curr_index):
    """返回 (added, modified, deleted)。

    added/modified 为 [(key, row), ...]；deleted 为 [key, ...]（清空表A只需主键）。
    prev_hashes 为 {key: hash}（上一轮持久化快照）。"""
    added, modified, deleted = [], [], []
    curr_hashes = {key: row_hash(row) for key, (row, _) in curr_index['rows'].items()}
    for key, (row, _row_no) in curr_index['rows'].items():
        if key not in prev_hashes:
            added.append((key, row))
        elif curr_hashes[key] != prev_hashes[key]:
            modified.append((key, row))
    for key in prev_hashes:
        if key not in curr_index['rows']:
            deleted.append(key)
    return added, modified, deleted


def build_delta(added, modified, deleted):
    """把三类变化合并为有序的 delta 列表（先增、再改、后删）。"""
    entries = []
    for key, row in added:
        entries.append(dict(op='add', pk=key, row=row))
    for key, row in modified:
        entries.append(dict(op='modify', pk=key, row=row))
    for key in deleted:
        entries.append(dict(op='delete', pk=key, row=None))
    return entries


def load_snapshot(root):
    """加载上一轮快照；返回 dict(meta, hashes={pk:hash}) 或 None。"""
    path = Path(root) / SNAPSHOT_FILE
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return None
    if not isinstance(data, dict) or not isinstance(data.get('rows'), dict):
        return None
    return data


def save_snapshot(root, index):
    """把本轮索引的哈希与元数据写为快照（不保存单元格内容）。"""
    payload = dict(version=1, captured_at=dt.datetime.now(dt.timezone.utc).isoformat(),
                   primary_key_column=index['primary_key_column'],
                   data_start_row=index['data_start_row'], range=index['range'],
                   width=index['width'], first_col=index['first_col'],
                   first_row=index['first_row'], header=index['header'],
                   rows=index_hashes(index))
    path = Path(root) / SNAPSHOT_FILE
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    Path(temp).replace(path)


def save_delta(root, delta, header, width):
    """把本轮变化写一份 JSON 记录，供排查与手动复查（上传以内存 delta 为准）。"""
    path = Path(root) / DELTA_FILE
    payload = dict(captured_at=dt.datetime.now(dt.timezone.utc).isoformat(),
                   width=width, header=header, delta=delta)
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    Path(temp).replace(path)

