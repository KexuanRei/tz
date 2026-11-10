"""Read per-cell styles and the primary-key index from 表C.xlsx.

表C is the downloaded source workbook (DATA/表C.xlsx). For every cell the sync
writes into 表A we must take the *format* of the cell in 表C at the same
primary-key row and column: font, fill color, border, alignment and number
format. Only the primary-key column's values are read, solely to map a primary
key to the row whose format should be reused; no other data is read or written.
"""
from pathlib import Path

from openpyxl import load_workbook

from sync_engine import parse_pk_column, data_start_row, pk_key


def table_c_path(root, config):
    """Absolute path of 表C.xlsx, or None when no format reference is configured."""
    value = (config['source'].get('format_file') or '').strip()
    if not value:
        return None
    return (Path(root) / value).resolve()


def _color_hex(color):
    """Return an RRGGBB hex string for an explicit RGB color, else None.

    Theme / indexed colors have no explicit RGB in the file; they are left
    alone rather than guessed, so the pasted cell keeps the sheet default.
    """
    if color is None or getattr(color, 'type', None) != 'rgb':
        return None
    try:
        value = color.rgb
    except Exception:
        return None
    if isinstance(value, str) and value:
        return value[-6:] if len(value) >= 6 else None
    return None


def _underline(font):
    value = getattr(font, 'underline', None)
    return value not in (None, 'none', '')


def extract_style(cell):
    """Convert one openpyxl cell into a plain style dict (primitives only).

    The dict uses flat keys so the clipboard writer can stay openpyxl-free.
    """
    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment
    style = dict(
        font_name=font.name,
        font_size=float(font.size) if font.size is not None else None,
        font_bold=bool(font.bold) if font.bold is not None else None,
        font_italic=bool(font.italic) if font.italic is not None else None,
        font_underline=_underline(font),
        font_strike=bool(font.strike) if font.strike is not None else None,
        font_color=_color_hex(font.color),
        fill_color=(_color_hex(fill.fgColor)
                    if fill is not None and getattr(fill, 'patternType', None) == 'solid'
                    else None),
        align_horizontal=alignment.horizontal if alignment is not None else None,
        align_vertical=alignment.vertical if alignment is not None else None,
        wrap_text=(bool(alignment.wrap_text)
                   if alignment is not None and alignment.wrap_text is not None else None),
        number_format=cell.number_format or 'General',
    )
    for side_name in ('top', 'right', 'bottom', 'left'):
        side = getattr(border, side_name, None)
        if side is not None and side.style:
            style['border_' + side_name] = side.style
            color = _color_hex(side.color)
            style['border_' + side_name + '_color'] = color or None
        else:
            style['border_' + side_name] = None
            style['border_' + side_name + '_color'] = None
    return style


class TableC:
    """Lazily loaded 表C: primary-key index plus per-cell format lookup."""

    def __init__(self, path, sheet_name, pk_number, data_start):
        self.path = path
        self.sheet_name = sheet_name
        self.pk_number = pk_number
        self.data_start = data_start
        self._book = None
        self._sheet = None
        self._pk_rows = None

    def _load(self):
        if self._book is not None:
            return
        book = load_workbook(self.path, data_only=True)
        try:
            if self.sheet_name:
                if self.sheet_name not in book.sheetnames:
                    available = '、'.join(book.sheetnames)
                    raise ValueError(
                        f'表C中找不到指定工作表“{self.sheet_name}”，请下载最新表C后重试'
                        + (f'（现有工作表：{available}）' if available else ''))
                sheet = book[self.sheet_name]
            else:
                sheet = book[book.sheetnames[0]]
        except Exception:
            book.close()
            raise
        self._book = book
        self._sheet = sheet
        self._build_index()

    def _build_index(self):
        index = {}
        for row in self._sheet.iter_rows(min_col=self.pk_number, max_col=self.pk_number,
                                         min_row=self.data_start, values_only=False):
            cell = row[0]
            key = pk_key(cell.value)
            if key is None:
                continue
            if key in index:
                raise ValueError(
                    f'表C主键列存在重复值“{key}”（第{index[key]}行与第{cell.row}行），'
                    f'请下载最新表C后再同步')
            index[key] = cell.row
        self._pk_rows = index

    def row_for(self, key):
        """表C row number for a primary key; raises when 表C lacks the key."""
        self._load()
        row = self._pk_rows.get(key)
        if row is None:
            raise ValueError(f'表C找不到主键“{key}”，请下载最新表C后再同步')
        return row

    def styles(self, row, columns):
        """Return one style dict per column for a given 表C row."""
        self._load()
        return [extract_style(self._sheet.cell(row, col)) for col in columns]

    def close(self):
        if self._book is not None:
            self._book.close()
            self._book = None
            self._sheet = None
            self._pk_rows = None


def load_table_c(root, config):
    """Return a TableC or None when 表C is not configured.

    Raises when the configured file is missing or not an .xlsx workbook, so a
    stale or deleted 表C stops the sync before any paste happens.
    """
    path = table_c_path(root, config)
    if path is None:
        return None
    if path.suffix.lower() != '.xlsx' or not path.is_file():
        raise ValueError('表C格式参考文件不存在或不是.xlsx文件，请检查路径或下载最新表C')
    _, pk_number = parse_pk_column(config)
    sheet_name = (config['source'].get('format_sheet') or '').strip()
    data_start = data_start_row(config)
    return TableC(path, sheet_name, pk_number, data_start)
