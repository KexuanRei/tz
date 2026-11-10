"""把下载的 表C 原始 xlsx 文件读成与整表复制同构的网格数据。

方法2：整表扫描不再依赖浏览器剪贴板（kdocs 对超大复制的富文本会截断），
而是复用“下载表C”得到的完整 xlsx（DATA/表C.xlsx，即表C 的原始快照），
用 openpyxl 直接读取数值并渲染成与网页显示一致的字符串。

数值渲染是本模块的核心：xlsx 存的是“原始值 + 数字格式”，而镜像与写回需要
“网页显示的字符串”。这里只覆盖表C 实际出现的格式（日期、整数、小数、百分比、
General），未知格式回退为 str()，保证绝不因格式解析失败而中断整轮。
"""
import datetime as dt
import re
from decimal import Decimal, ROUND_HALF_UP, localcontext
from pathlib import Path

from openpyxl import load_workbook

from sync_engine import start_cell
from sync_columns import effective_skip_columns
from update_excel import address


def _positive_section(number_format):
    """取数字格式的正数段（第一个 ';' 之前），去掉颜色/条件与对齐占位。"""
    section = str(number_format or '').split(';', 1)[0]
    section = re.sub(r'\[[^\]]*\]', '', section)          # [Red]、[$-409]、[h] 等标记
    section = re.sub(r'[_*](\\.|.)', '', section)          # _x / *x 对齐占位
    return section.replace('"', '')


def _decimal_places(section):
    if '.' not in section:
        return 0
    fraction = section.split('.', 1)[1]
    count = 0
    for ch in fraction:
        if ch in '0#?':
            count += 1
        elif ch in 'eE':
            break
        # 逗号、百分号等分隔符不计小数位
    return count


def _general_number(value):
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, int):
        return str(value)
    number = float(value)
    if number == int(number) and abs(number) < 1e15:
        return str(int(number))
    return repr(number)


def _render_number(value, section):
    if section in ('', 'General', 'general'):
        return _general_number(value)
    is_percent = '%' in section
    decimals = _decimal_places(section)
    # 用 Decimal 按“四舍五入（half-up）”渲染，与 WPS/Excel 的显示一致。
    # Python 的 '{:.Nf}'.format 走“银行家舍入”且受浮点表示影响，会把 38.05 显示成 38.0，
    # 而 WPS 显示 38.1；这会让线下对比（表B vs 表C）产生 0.1 的假差异，进而误粘贴。
    number = Decimal(str(value))
    if is_percent:
        number *= 100
    quantum = Decimal('1') if decimals == 0 else Decimal('1e-{}'.format(decimals))
    with localcontext() as context:
        context.prec = max(28, number.adjusted() + decimals + 2)
        text = str(number.quantize(quantum, rounding=ROUND_HALF_UP))
    return text + ('%' if is_percent else '')


def _render_date(value, number_format):
    """把 Excel 日期格式（yyyy/m/d 等）渲染成显示字符串。"""
    section = str(number_format or '').split(';', 1)[0]
    out = []
    i = 0
    length = len(section)
    while i < length:
        ch = section[i]
        if ch in 'ymdhs':
            j = i
            while j < length and section[j] == ch:
                j += 1
            count = j - i
            if ch == 'y':
                out.append(str(value.year).zfill(2) if count == 2 else str(value.year))
            elif ch == 'm':
                out.append(str(value.month) if count == 1 else str(value.month).zfill(2))
            elif ch == 'd':
                out.append(str(value.day) if count == 1 else str(value.day).zfill(2))
            elif ch == 'h':
                out.append(str(value.hour) if count == 1 else str(value.hour).zfill(2))
            elif ch == 's':
                out.append(str(value.second) if count == 1 else str(value.second).zfill(2))
            i = j
        elif ch == '"':
            j = section.index('"', i + 1) if '"' in section[i + 1:] else length
            out.append(section[i + 1:j])
            i = j + 1
        elif ch == '\\':
            if i + 1 < length:
                out.append(section[i + 1])
                i += 2
            else:
                i += 1
        else:
            out.append(ch)
            i += 1
    return ''.join(out)


def render_value(value, number_format=None):
    """把一个 xlsx 单元格值渲染成与网页显示一致的字符串。"""
    if value is None:
        return ''
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, (dt.datetime, dt.date)):
        return _render_date(value, number_format)
    if isinstance(value, dt.time):
        return value.isoformat()
    if isinstance(value, (int, float, Decimal)):
        return _render_number(value, _positive_section(number_format))
    return str(value)


def _data_path(root, config):
    value = (config.get('source', {}).get('format_file') or '').strip() or 'DATA/表C.xlsx'
    return (Path(root) / value).resolve()


def _sheet_name(config, workbook):
    name = (config.get('source', {}).get('format_sheet') or '').strip() \
        or (config.get('source', {}).get('sheet_name') or '').strip()
    return name if name in workbook.sheetnames else workbook.sheetnames[0]


def read_table_data(root, config):
    """读取 DATA/表C.xlsx（表C 快照）的数据区域，返回与整表复制同构的网格。"""
    path = _data_path(root, config)
    if not path.is_file():
        raise ValueError('表C 数据文件不存在：' + str(path) + '；请先下载表C（表C 原始 Excel）')
    _start_name, start_col, start_row = start_cell(config)

    # data_only=False 让公式单元格返回公式表达式本身（data_type=='f'），而非计算结果，
    # 从而把公式按公式同步到表A；普通单元格与 data_only=True 的原始值一致。
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        sheet = workbook[_sheet_name(config, workbook)]
        last_col = sheet.max_column
        last_row = sheet.max_row
        # 表C 下载文件里可能残留“幽灵列”：仅含格式/空白、但把 max_column 撑大的尾列。
        # 这类列没有数据，却会让镜像“结束列”漂移，导致身份校验失败。配置的 scan_range
        # 是数据区域的权威声明，因此把结束列收敛到 scan_range 的右边界（若声明了范围）。
        scan_range = (config.get('source', {}).get('scan_range') or '').strip()
        if scan_range and ':' in scan_range:
            try:
                _scan_left, _scan_top = address(scan_range.split(':', 1)[0])
                scan_right, _scan_bottom = address(scan_range.split(':', 1)[1])
                if scan_right < last_col:
                    last_col = scan_right
            except ValueError:
                pass
        if last_row < start_row or last_col < start_col:
            raise ValueError('表C 数据文件没有数据：起始单元格之后没有数据行（' + str(path) + '）')
        skipped = set(effective_skip_columns(config, start_col, last_col))
        width = last_col - start_col + 1
        total_rows = last_row - start_row + 1
        rows = [[''] * width for _ in range(total_rows)]
        for ri, row in enumerate(sheet.iter_rows(min_row=start_row, max_row=last_row,
                                                 min_col=start_col, max_col=last_col,
                                                 values_only=False)):
            for ci, cell in enumerate(row, start=start_col):
                if ci in skipped:
                    continue
                rows[ri][ci - start_col] = render_value(getattr(cell, 'value', None),
                                                        getattr(cell, 'number_format', 'General'))
    finally:
        workbook.close()

    merges = set()
    try:
        from format_reference import load as load_layout
        layout = load_layout(root, config)
        for c1, r1, c2, r2 in (layout or {}).get('merges', ()):
            if r1 >= start_row and c1 >= start_col:
                merges.add((c1 - start_col, r1 - start_row, c2 - start_col, r2 - start_row))
    except Exception:
        merges = set()
    return dict(rows=rows, first_col=start_col, first_row=start_row,
                width=width, row_count=total_rows, last_row=last_row, last_col=last_col,
                merges=merges)
