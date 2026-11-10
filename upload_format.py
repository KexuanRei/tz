"""Build the upload sheet from A1 values and C formatting, without saving either.

Only Excel's formats paste is used for C: template values, formulas and links are
never read or pasted. The caller owns the Excel application and the online UI.
"""
from contextlib import contextmanager
from dataclasses import dataclass
from math import isfinite
from pathlib import Path

from copy_range import column_name


XL_PASTE_VALUES = -4163
XL_PASTE_FORMATS = -4122
XL_WORKSHEET = -4167
MAX_CHECK_CELLS = 100_000


@dataclass(frozen=True)
class PreparedUpload:
    sheet: object
    layout: dict


def _positive_dimension(value, label):
    if isinstance(value, bool):
        raise ValueError('表C的%s无效' % label)
    try:
        value = float(value)
    except (TypeError, ValueError):
        raise ValueError('表C的%s无效' % label) from None
    if not isfinite(value) or value <= 0:
        raise ValueError('表C的%s必须是正数' % label)
    return value


def _runs(size, overrides, default, label):
    """Return inclusive contiguous runs, including C's default dimensions."""
    result = []
    for index in range(1, size + 1):
        value = _positive_dimension(overrides.get(index, default), label)
        if result and result[-1][2] == value:
            result[-1] = (result[-1][0], index, value)
        else:
            result.append((index, index, value))
    return result


def layout_metadata(reference, rows, cols, default_row_height=None,
                    default_col_width=None, source_left=1, source_top=1):
    """Normalize C layout within A1's upload extent.

    Table C dimensions and merges use the ORIGINAL sheet coordinates; the
    uploaded scratch uses 1-based coordinates. ``source_left``/``source_top``
    translate between them (both 1 for the default A1 start).

    A merge partly outside the upload cannot be clipped without changing its
    meaning. Reject it before copying rather than silently cutting the merge.
    """
    if (isinstance(rows, bool) or isinstance(cols, bool)
            or not isinstance(rows, int) or not isinstance(cols, int)
            or not 1 <= rows <= 1048576 or not 1 <= cols <= 16384):
        raise ValueError('上传范围超出Excel行列限制')
    for value, label, maximum in ((source_left, '源起始列', 16384), (source_top, '源起始行', 1048576)):
        if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= maximum:
            raise ValueError(f'{label}必须是1至{maximum}的整数')
    if reference is None:
        raise ValueError('请先设置表C格式来源')
    height = reference.get('default_row_height')
    width = reference.get('default_col_width')
    if height is None:
        height = default_row_height
    if width is None:
        width = default_col_width
    height = _positive_dimension(height, '默认行高')
    width = _positive_dimension(width, '默认列宽')
    last_col = source_left + cols - 1
    last_row = source_top + rows - 1
    merges = []
    absolute_merges = []
    for merge in reference.get('merges', []):
        if (len(merge) != 4 or any(isinstance(v, bool) or not isinstance(v, int) for v in merge)
                or not 1 <= merge[0] <= merge[2] <= 16384
                or not 1 <= merge[1] <= merge[3] <= 1048576):
            raise ValueError('表C包含无效合并区域')
        left, top, right, bottom = merge
        if right < source_left or left > last_col or bottom < source_top or top > last_row:
            continue
        if left < source_left or right > last_col or top < source_top or bottom > last_row:
            raise ValueError('上传范围切断表C合并区域%s，请扩大表A1上传范围或调整表C合并' %
                             _address(left, top, right, bottom))
        internal = (left - source_left + 1, top - source_top + 1,
                    right - source_left + 1, bottom - source_top + 1)
        for other in merges:
            if not (internal[2] < other[0] or internal[0] > other[2]
                    or internal[3] < other[1] or internal[1] > other[3]):
                raise ValueError('表C包含重叠合并区域')
        merges.append(internal)
        absolute_merges.append(merge)
    column_widths = {int(key) - source_left + 1: value
                     for key, value in reference.get('column_widths', {}).items()
                     if source_left <= int(key) <= last_col}
    row_heights = {int(key) - source_top + 1: value
                   for key, value in reference.get('row_heights', {}).items()
                   if source_top <= int(key) <= last_row}
    return dict(sheet=reference.get('sheet'), rows=rows, columns=cols, merges=merges,
                absolute_merges=absolute_merges,
                default_row_height=height, default_col_width=width,
                row_heights=row_heights, column_widths=column_widths,
                row_runs=_runs(rows, row_heights, height, '行高'),
                column_runs=_runs(cols, column_widths, width, '列宽'),
                pane=reference.get('pane'))


def _address(left, top, right, bottom):
    return '%s%d:%s%d' % (column_name(left), top, column_name(right), bottom)


def _matrix(value, rows, cols):
    # COM returns a scalar for one cell, and a tuple of row tuples otherwise.
    if rows == 1 and cols == 1:
        return ((value,),)
    if (not isinstance(value, (tuple, list)) or len(value) != rows
            or any(not isinstance(row, (tuple, list)) or len(row) != cols for row in value)):
        raise ValueError('无法完整读取表A1合并区域中的数据，已停止上传')
    return value


def validate_merges(source_sheet, merges):
    """Check A1 values before formats-paste can merge any cells in scratch.

    Bulk COM reads are bounded for large merged rectangles. The exception only
    contains coordinates, never the user's cell contents.
    """
    for left, top, right, bottom in merges:
        width = right - left + 1
        chunk_rows = max(1, MAX_CHECK_CELLS // width)
        for row in range(top, bottom + 1, chunk_rows):
            end = min(bottom, row + chunk_rows - 1)
            values = source_sheet.Range(_address(left, row, right, end)).Value2
            for row_offset, values_row in enumerate(_matrix(values, end - row + 1, width)):
                for col_offset, value in enumerate(values_row):
                    if row + row_offset == top and col_offset == 0:
                        continue
                    if value is not None and value != '':
                        conflict = '%s%d' % (column_name(left + col_offset), row + row_offset)
                        raise ValueError('表C合并区域%s会覆盖表A1的%s数据；请先调整表C合并，表C尚未修改' %
                                         (_address(left, top, right, bottom), conflict))


def _sheet(book, name, label):
    try:
        return book.Worksheets(name) if name else book.Worksheets(1)
    except Exception:
        raise ValueError('%s中找不到工作表“%s”' % (label, name)) from None


def _apply_layout(sheet, layout):
    for start, end, value in layout['column_runs']:
        sheet.Range('%s:%s' % (column_name(start), column_name(end))).ColumnWidth = value
    for start, end, value in layout['row_runs']:
        sheet.Range('%d:%d' % (start, end)).RowHeight = value
    for left, top, right, bottom in layout['merges']:
        sheet.Range(_address(left, top, right, bottom)).Merge()


def _match_column_widths(sheet, width_points):
    """Keep physical width when C and the blank book use different Normal fonts.

    Excel's ColumnWidth is measured in Normal-font characters; Width is points
    and read-only. A bounded correction per column avoids changing the source
    template or assuming that both workbooks have the same font metrics.
    """
    for number, target in width_points.items():
        column = sheet.Columns(number)
        for attempt in range(9):
            actual = _positive_dimension(column.Width, '临时工作簿列宽')
            if abs(actual - target) <= 0.75:  # Excel rounds at screen-pixel granularity.
                break
            if attempt == 8:
                raise ValueError('临时工作簿的%s列宽无法匹配表C，已停止上传' % column_name(number))
            current = _positive_dimension(column.ColumnWidth, '临时工作簿列宽')
            corrected = max(0.01, min(255.0, current * target / actual))
            column.ColumnWidth = corrected


def _copy_normal_font(template, scratch):
    """The blank book's width ruler must use C's Normal font metrics."""
    source_font = template.Styles('Normal').Font
    target_font = scratch.Styles('Normal').Font
    target_font.ThemeFont = 0
    for attribute in ('Name', 'Size', 'Bold', 'Italic'):
        setattr(target_font, attribute, getattr(source_font, attribute))


def _column_points(template_sheet, scratch_sheet, cols, source_left=1):
    points = {}
    for col in range(1, cols + 1):
        raw_width = float(template_sheet.Columns(source_left + col - 1).Width)
        if raw_width == 0:
            # Hidden C columns report Width == 0 although the xlsx keeps their
            # positive, expanded ColumnWidth. The scratch column has that saved
            # width and C's Normal font: measure it there without unhiding C.
            probe = scratch_sheet.Columns(col)
            probe.Hidden = False
            raw_width = probe.Width
        points[col] = _positive_dimension(raw_width, '列宽')
    return points


@contextmanager
def prepared_upload(excel, a1_path, a1_sheet, format_path, format_sheet,
                    rows, cols, reference, *, source_left=1, source_top=1):
    """Yield PreparedUpload(sheet, layout); close all owned books on any exit.

    The caller must use its own Excel COM instance with macros/events disabled.
    The context opens A1 and C read-only and never saves a workbook. Only the
    yielded scratch worksheet may be copied into the online destination.
    """
    if not format_path or reference is None:
        raise ValueError('请先设置表C格式来源')
    a1_path, format_path = Path(a1_path).resolve(), Path(format_path).resolve()
    if a1_path == format_path:
        raise ValueError('表C格式来源不能是表A1本身')
    books = []
    try:
        a1 = excel.Workbooks.Open(str(a1_path), UpdateLinks=0, ReadOnly=True,
                                 IgnoreReadOnlyRecommended=True, AddToMru=False)
        books.append(a1)
        source = _sheet(a1, a1_sheet, '表A1')
        template = excel.Workbooks.Open(str(format_path), UpdateLinks=0, ReadOnly=True,
                                       IgnoreReadOnlyRecommended=True, AddToMru=False)
        books.append(template)
        template_sheet = _sheet(template, format_sheet or reference.get('sheet'), '表C')
        layout = layout_metadata(reference, rows, cols,
                                 default_row_height=template_sheet.StandardHeight,
                                 default_col_width=template_sheet.StandardWidth,
                                 source_left=source_left, source_top=source_top)
        validate_merges(source, layout['absolute_merges'])
        scratch = excel.Workbooks.Add(XL_WORKSHEET)
        books.append(scratch)
        _copy_normal_font(template, scratch)
        sheet = scratch.Worksheets(1)
        extent = _address(1, 1, cols, rows)
        source_extent = _address(source_left, source_top, source_left + cols - 1, source_top + rows - 1)
        # Native bulk paste keeps literal strings literal and turns formulas
        # into their current A1 values; C never contributes data.
        from retry_com import retry_call
        retry_call(lambda: source.Range(source_extent).Copy())
        retry_call(lambda: sheet.Range(extent).PasteSpecial(Paste=XL_PASTE_VALUES, SkipBlanks=False))
        sheet.Range(extent).UnMerge()
        retry_call(lambda: template_sheet.Range(source_extent).Copy())
        retry_call(lambda: sheet.Range(extent).PasteSpecial(Paste=XL_PASTE_FORMATS, SkipBlanks=False))
        # Explicitly reapply C's structure, including the empty-merges case.
        sheet.Range(extent).UnMerge()
        _apply_layout(sheet, layout)
        # Width is Excel's measured physical width in points, unlike
        # ColumnWidth's font-dependent character units. Online controls using
        # pixels can convert points by 96 / 72 without guessing glyph widths.
        layout['column_width_points'] = _column_points(template_sheet, sheet, cols, source_left)
        _match_column_widths(sheet, layout['column_width_points'])
        excel.CutCopyMode = False
        yield PreparedUpload(sheet, layout)
    finally:
        for book in reversed(books):
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
