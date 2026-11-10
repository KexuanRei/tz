"""Read user-pasted spreadsheet TSV. Never contacts WPS or reads the clipboard itself."""
import re

from update_excel import address
from read_limits import MAX_CELLS, CellLimitError

MAX_CELL_TEXT = 32767
_CONTROL = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')


class PasteShapeError(ValueError):
    """A parsed field cannot be stored in a cell (too long or control characters)."""

    def __init__(self, message, *, row, column, length, control=None):
        super().__init__(message)
        self.diagnostic_context = dict(row=row, column=column, length=length,
                                       control=control)


def _clean(value):
    """Excel cannot store these characters; drop them and cap the length."""
    return _CONTROL.sub('', value)[:MAX_CELL_TEXT]


def _quoted_field(text, index, length):
    """Parse a field that starts with a quote; None when it is not properly closed.

    Excel quotes cells that contain tabs/newlines/quotes and doubles inner quotes.
    WPS sometimes writes a lone leading quote or an unescaped inner quote.  A quote
    is also refused when the field would grow past the Excel cell limit, so a stray
    quote can never swallow the following rows into one giant cell."""
    position = index + 1
    parts = []
    while position < length:
        character = text[position]
        if character == '"':
            if position + 1 < length and text[position + 1] == '"':
                parts.append('"')
                position += 2
                continue
            following = text[position + 1] if position + 1 < length else ''
            if following in ('', '\t', '\r', '\n'):
                return ''.join(parts), position + 1
            parts.append('"')
            position += 1
            continue
        parts.append(character)
        position += 1
        if position - index > MAX_CELL_TEXT:
            return None
    return None


def _field(text, index, length):
    if index < length and text[index] == '"':
        quoted = _quoted_field(text, index, length)
        if quoted is not None:
            return quoted
    position = index
    while position < length and text[position] not in ('\t', '\r', '\n'):
        position += 1
    return text[index:position], position


def read_rows(text, *, quoted=True):
    """Tolerant TSV reader: quoted cells may hold tabs/newlines, stray quotes stay literal."""
    rows, row = [], []
    index, length = 0, len(text)
    while index < length:
        if quoted:
            value, index = _field(text, index, length)
        else:
            start = index
            while index < length and text[index] not in ('\t', '\r', '\n'):
                index += 1
            value = text[start:index]
        row.append(value)
        if index >= length:
            break
        character = text[index]
        if character == '\t':
            index += 1
            if index >= length:  # 行尾制表符表示后面还有一个空单元格
                row.append('')
            continue
        if character == '\r':
            index += 1
            if index < length and text[index] == '\n':
                index += 1
        elif character == '\n':
            index += 1
        rows.append(row)
        row = []
    if row:
        rows.append(row)
    return rows


def parse(text, start_cell, max_cells, *, retain_rows=True, allow_blank=False, literal=False):
    if not isinstance(text, str) or (not allow_blank and not text.strip()):
        raise ValueError('请先粘贴从表A复制的单元格数据')
    first_col, first_row = address(start_cell)
    rows, width, height, header = [], 0, 0, []
    nonempty_count = 0
    exceeded = False
    # 每行最后一个有文字的列、每一列是否出现过文字：只占 O(行+列)，不保存单元格内容，
    # 用来确认“富文本少掉的末尾列在纯文本里确实是空白”，而不是复制掉了数据。
    row_ends = []
    column_nonempty = bytearray()
    column_counts = []
    for raw in read_rows(text if text else '\n', quoted=not literal):
        row = list(raw) or ['']
        if literal:
            # 字面解析：整行按制表符切分，控制字符删除、超长截断，绝不让整批失败。
            row = [_clean(value) for value in row]
        else:
            for column, value in enumerate(row, 1):
                match = _CONTROL.search(value)
                if len(value) > MAX_CELL_TEXT or match:
                    raise PasteShapeError('单元格文本过长或包含不支持的控制字符',
                                          row=height + 1, column=column, length=len(value),
                                          control=hex(ord(match.group())) if match else None)
        end = -1
        for index, value in enumerate(row):
            if value.strip():
                end = index
                if index >= len(column_nonempty):
                    column_nonempty.extend(bytes(index + 1 - len(column_nonempty)))
                column_nonempty[index] = 1
                if index >= len(column_counts):
                    column_counts.extend([0] * (index + 1 - len(column_counts)))
                column_counts[index] += 1
        row_ends.append(end)
        width = max(width, len(row))
        height += 1
        nonempty_count += sum(value != '' for value in row)
        if height == 1: header = row
        if width * height > MAX_CELLS:
            raise CellLimitError(height, width, min(max_cells, MAX_CELLS), '粘贴内容', exact=False)
        if width * height > max_cells:
            exceeded = True
            rows.clear()
        if retain_rows and not exceeded: rows.append(row)
        if first_row + height - 1 > 1048576 or first_col + width - 1 > 16384:
            raise ValueError('粘贴内容超出Excel行列范围')
    if not width or not height:
        raise ValueError('没有可导入的数据')
    if exceeded:
        # Finish scanning without retaining values, so the suggested limit covers
        # the complete rectangle instead of the first row that exceeds the limit.
        raise CellLimitError(height, width, max_cells, '粘贴内容')
    for row in rows: row.extend([''] * (width - len(row)))
    return dict(first_column=first_col, first_row=first_row, width=width, rows=rows,
                row_count=height, header=header, nonempty_count=nonempty_count,
                row_ends=row_ends, column_nonempty=bytes(column_nonempty),
                column_counts=column_counts)
