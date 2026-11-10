"""Parse a user-selected table C destination without expanding its boundaries."""
from dataclasses import dataclass
import re

from copy_range import column_name
from update_excel import address


class UploadTargetError(ValueError):
    def __init__(self, code, message, **context):
        super().__init__(message)
        self.code = code
        self.diagnostic_context = context


@dataclass(frozen=True)
class Destination:
    ref: str
    column: int
    row: int
    last_column: int | None = None
    last_row: int | None = None

    @property
    def start(self):
        return f'{column_name(self.column)}{self.row}'

    @property
    def bounded(self):
        return self.last_column is not None

    @property
    def columns(self):
        return self.last_column - self.column + 1 if self.bounded else None

    @property
    def rows(self):
        return self.last_row - self.row + 1 if self.bounded else None


def parse_destination(value):
    hint = '表C目标位置请填写单个起点（如A1）或完整区域（如A1:BW60）'
    if not isinstance(value, str):
        raise UploadTargetError('UPLOAD_TARGET_INVALID', hint)
    ref = value.strip().upper()
    if not re.fullmatch(r'\$?[A-Z]{1,3}\$?[1-9][0-9]*(?:\s*:\s*\$?[A-Z]{1,3}\$?[1-9][0-9]*)?', ref):
        raise UploadTargetError('UPLOAD_TARGET_INVALID', hint)
    ref = ref.replace('$', '').replace(' ', '')
    parts = [part.strip() for part in ref.split(':')]
    try:
        col, row = address(parts[0])
        end_col, end_row = address(parts[-1])
    except ValueError as exc:
        raise UploadTargetError('UPLOAD_TARGET_INVALID', '表C目标位置超出Excel范围：最多XFD列、1048576行', destination=ref) from exc
    if end_col < col or end_row < row:
        raise UploadTargetError('UPLOAD_TARGET_INVALID', '表C目标区域的结束位置不能在起始位置之前', destination=ref)
    ref = ':'.join(parts)
    if len(parts) == 1:
        return Destination(ref, col, row)
    return Destination(ref, col, row, end_col, end_row)


def plan_upload(rows, columns, value, overflow='stop', source_range=None, allow_blank=False):
    """Return the source prefix and actual target; never stretch or repeat cells.

    source_range: 可选的表A1源区域（如 A90:BW100 或单个起点 A90）。
    默认从A1开始取到表A1已使用范围的末尾。
    allow_blank: 源区域整段在表A1数据范围之外（全是空白）时，允许按请求的矩形
    在表C上“空白覆盖”，即清空表C同位置的那一段。"""
    if overflow not in ('stop', 'clip'):
        raise UploadTargetError('UPLOAD_TARGET_INVALID', '上传范围不足时的处理方式无效')
    if type(rows) is not int or type(columns) is not int or rows < 1 or columns < 1:
        raise UploadTargetError('UPLOAD_TARGET_INVALID', '表A1没有有效上传范围')
    target = parse_destination(value)
    source_rows, source_columns = rows, columns
    source_left, source_top = 1, 1
    blank_overwrite = False
    if source_range:
        src = parse_destination(source_range)
        source_left, source_top = src.column, src.row
        available_rows = max(0, rows - source_top + 1)
        available_columns = max(0, columns - source_left + 1)
        requested_rows = src.rows if src.bounded else available_rows
        requested_columns = src.columns if src.bounded else available_columns
        rows = min(requested_rows, available_rows)
        columns = min(requested_columns, available_columns)
        if rows < 1 or columns < 1:
            if allow_blank and src.bounded:
                # 空白覆盖：源区域整段在数据范围之外（全是空白），按请求的矩形
                # 在表C同位置覆盖成空白（等于清空那一段）。必须显式开启才允许。
                source_left, source_top = src.column, src.row
                rows = min(src.rows, 1048576 - source_top + 1)
                columns = min(src.columns, 16384 - source_left + 1)
                if rows < 1 or columns < 1:
                    raise UploadTargetError('UPLOAD_TARGET_INVALID',
                        '空白覆盖范围超出Excel行列限制', source_range=source_range)
                blank_overwrite = True
            else:
                data_range = f'A1:{column_name(source_columns)}{source_rows}'
                raise UploadTargetError('UPLOAD_TARGET_INVALID',
                    '表A1源区域%s与表A1数据范围（%s行 × %s列，即 %s）没有交集：'
                    '第%s行以后表A1没有数据。源区域留空表示整表上传；只想传某一段就填实际有数据的范围'
                    '（例如 %s），别把“表C目标位置”误填成源区域。'
                    '如果就是要用空白覆盖、清空表C那一段，请勾选“允许空白覆盖”。尚未向表C粘贴数据。'
                    % (source_range, source_rows, source_columns, data_range,
                       source_top, data_range), source_range=source_range)
        truncated_source = rows < requested_rows or columns < requested_columns
    else:
        truncated_source = False
    if target.bounded:
        if overflow == 'stop' and (rows > target.rows or columns > target.columns):
            raise UploadTargetError('UPLOAD_TARGET_TOO_SMALL',
                f'表A1本次上传范围为{rows}行 × {columns}列，表C目标区域{target.ref}只能容纳'
                f'{target.rows}行 × {target.columns}列。请扩大目标区域，或选择“只取表A1左上角能放下的部分”；尚未向表C粘贴数据。',
                requested_destination=target.ref, source_rows=rows, source_columns=columns,
                available_rows=target.rows, available_columns=target.columns)
        rows, columns = min(rows, target.rows), min(columns, target.columns)
    end_col, end_row = target.column + columns - 1, target.row + rows - 1
    if end_col > 16384 or end_row > 1048576:
        raise UploadTargetError('UPLOAD_TARGET_INVALID', '表C实际写入范围超出Excel行列限制', destination=target.ref)
    return dict(rows=rows, columns=columns,
                source_range=f'{column_name(source_left)}{source_top}:{column_name(source_left+columns-1)}{source_top+rows-1}',
                source_left=source_left, source_top=source_top,
                destination=target.start, requested_destination=target.ref, bounded=target.bounded,
                target_range=f'{target.start}:{column_name(end_col)}{end_row}',
                truncated=False if blank_overwrite
                          else (truncated_source or rows < source_rows or columns < source_columns),
                blank_overwrite=blank_overwrite,
                source_rows=source_rows, source_columns=source_columns)
