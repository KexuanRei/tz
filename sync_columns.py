"""Explicit sync exclusions and inclusions. Columns retain their original worksheet coordinates.

支持两种列过滤，二者可同时使用：
- 跳过列（skip_columns）：黑名单，这些列完全不读取/比较/粘贴；
- 仅同步列（sync_columns）：白名单，只读取/比较/粘贴这些列（主键列始终保留），其余列完全不关心。
"""
import re

from copy_range import column_name
from update_excel import column


def _parse_column_list(value, label, all_hint):
    if not isinstance(value, str):
        raise ValueError(f'{label}须填写列字母，例如 F、F,H 或 F:H；{all_hint}')
    if not value.strip():
        return ()
    columns = set()
    for part in re.split(r'[,，;；\s]+', value.strip().upper()):
        if not part:
            continue
        match = re.fullmatch(r'([A-Z]{1,3})(?:[:-]([A-Z]{1,3}))?', part)
        if not match:
            raise ValueError(f'{label}格式错误，请填写 F、F,H,K 或 F:H，不要填写行号')
        try:
            start = column(match[1])
            end = column(match[2] or match[1])
        except ValueError as exc:
            raise ValueError(f'{label}须在 A 至 XFD 范围内') from exc
        if end < start:
            raise ValueError(f'{label}范围须从左到右，例如 F:H')
        columns.update(range(start, end + 1))
    if not columns:
        raise ValueError(f'{label}须填写列字母；{all_hint}')
    return tuple(sorted(columns))


def parse_skip_columns(value):
    return _parse_column_list(value, '跳过列', '扫描全部列时留空')


def parse_sync_columns(value):
    """解析「仅同步列」白名单；空字符串 = 同步全部列。"""
    return _parse_column_list(value, '仅同步列', '留空表示同步全部列')


def kept_spans(first, last, skipped):
    skipped = set(skipped)
    start = None
    for c in range(first, last + 1):
        if c in skipped:
            if start is not None:
                yield start, c - 1
                start = None
        elif start is None:
            start = c
    if start is not None:
        yield start, last


def sync_skip_columns(config):
    """黑名单「跳过列」，返回绝对列号元组（主键列不能跳过）。"""
    columns = parse_skip_columns(config.get('sync', {}).get('skip_columns', ''))
    pk = str(config.get('sync', {}).get('primary_key_column') or 'D').strip().upper()
    if column(pk) in columns:
        raise ValueError(f'主键列 {pk} 不能跳过，否则无法匹配对应行')
    return columns


def effective_skip_columns(config, first_col, last_col):
    """返回实际跳过的绝对列号元组（黑名单 ∪ 白名单之外的列，主键列始终保留）。

    显式跳过列始终生效；若配置了「仅同步列」，则 [first_col..last_col] 内除
    白名单列和主键列之外的列都视为跳过（完全不参与读取/比较/粘贴）。主键列
    必须保留用于行匹配。
    """
    skip = set(sync_skip_columns(config))
    only = parse_sync_columns(config.get('sync', {}).get('sync_columns', ''))
    if only:
        pk_name = str(config.get('sync', {}).get('primary_key_column') or 'D').strip().upper()
        pk = column(pk_name)
        allow = set(only) | {pk}
        for col in range(first_col, last_col + 1):
            if col not in allow:
                skip.add(col)
    return tuple(sorted(skip))


def format_skip_columns(columns):
    groups = []
    for value in sorted(set(columns)):
        if groups and value == groups[-1][1] + 1:
            groups[-1][1] = value
        else:
            groups.append([value, value])
    return ','.join(column_name(a) if a == b else f'{column_name(a)}:{column_name(b)}'
                    for a, b in groups)
