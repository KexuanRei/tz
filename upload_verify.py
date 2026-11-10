"""Verify each online upload using normal rich clipboard read-back, without data logs."""
import re
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext

from lxml import etree

from clipboard_html import table_fragment
from excel_clipboard import _TableShape, _check_html_errors, inspect_copy
from update_excel import address
from copy_range import column_name


class CellResult(str):
    """保留原文本用于主键定位，另存数值/公式结果用于内容比较。"""
    def __new__(cls, text, result):
        item = super().__new__(cls, text)
        item.result = result
        return item

    def __getnewargs__(self):
        return str(self), self.result


def comparison_result(value):
    return value.result if isinstance(value, CellResult) else value


class _MergedShape(_TableShape):
    def __init__(self, max_cells, rows=None):
        super().__init__(max_cells)
        self.merges = []
        self.rows = rows
        self.active_cell = None

    def start(self, tag, attrs):
        super().start(tag, attrs)
        if tag in ('td', 'th') and self.in_table:
            self.active_cell = (self.cell_row, self.cell_col, dict(attrs), [])
            cs, rs = int(attrs.get('colspan', '1')), int(attrs.get('rowspan', '1'))
            if cs > 1 or rs > 1:
                self.merges.append((self.col - cs, self.row, self.col - 1, self.row + rs - 1))
        elif tag == 'br' and self.active_cell is not None:
            self.active_cell[3].append('\n')

    def data(self, text):
        super().data(text)
        if self.active_cell is not None:
            self.active_cell[3].append(text)

    def end(self, tag):
        if tag in ('td', 'th') and self.active_cell is not None:
            row, col, attrs, parts = self.active_cell
            self.active_cell = None
            if self.rows is not None and row < len(self.rows) and col < len(self.rows[row]):
                plain = self.rows[row][col]
                display = ''.join(parts)
                cached = attrs.get('x:num')
                formula = attrs.get('x:fmla') or attrs.get('data-sheets-formula')
                result = None
                if formula:
                    # 公式单元格：以公式表达式本身作为比较身份（不解析、不执行），
                    # 使「复制公式本身」能按公式文本与表A 的公式对比，而非按计算结果。
                    result = formula
                elif cached is not None:
                    try:
                        number = Decimal(cached.strip() or display.strip())
                        if number.is_finite():
                            result = number
                    except InvalidOperation:
                        result = _numeric_result(display) if not cached.strip() else None
                if result is None and ('x:str' not in attrs and normalize_text(plain).startswith('=')):
                    # 没有 x:fmla 却以 '=' 开头的纯文本：无法确认计算结果，拒绝把公式文本当值粘贴。
                    candidate = plain if not normalize_text(plain).startswith('=') else display
                    if normalize_text(candidate).startswith('='):
                        raise ValueError('剪贴板仅提供公式，未提供计算结果；已停止，未将公式文本作为差异粘贴')
                    if candidate:
                        result = candidate
                    else:
                        raise ValueError('无法确认公式计算结果，已停止本次读取')
                if result is not None:
                    self.rows[row][col] = CellResult(plain, result)
        super().end(tag)


def snapshot(clip, ref, max_cells, *, strict_rows=False, pad_trailing_columns=False,
             pad_trailing_rows=False):
    """Keep relative merge positions, allowing comparison at a shifted destination."""
    config = {'source': {'max_cells': max_cells}, 'export': {'target': {'mode': 'cells'}}}
    origin = ref.split(':')[0]
    grid = inspect_copy(clip, config, origin, 'selection', expected_range=ref, batch=True,
                        strict_rows=strict_rows, pad_trailing_columns=pad_trailing_columns,
                        pad_trailing_rows=pad_trailing_rows)
    target = _MergedShape(max_cells, grid['rows'])
    parser = etree.HTMLParser(target=target, no_network=True)
    html = table_fragment(clip['html'])
    for offset in range(0, len(html), 65536):
        parser.feed(html[offset:offset + 65536])
        _check_html_errors(parser)
    parser.close()
    _check_html_errors(parser)
    return dict(rows=grid['rows'], row_count=grid['row_count'], width=grid['width'],
                merges=set(target.merges), unverified_rows=grid.get('unverified_rows', 0))


_SPACES = dict.fromkeys(map(ord, '\u00a0\u3000\u200b\ufeff'), None)

# 与 README 的“忽略首尾空格、特殊空白”一致：strip() 只去除 ASCII 空白，
# 窄空格（U+2009）、窄不换行空格（U+202F）、图空（U+2007）等 Unicode 空白仍会
# 造成“看似没差异却判定为不同”。这里按 Unicode 空白定义去除每行首尾的空白。
_WS_EDGE = re.compile(r'^\s+|\s+$')


def normalize_text(value):
    """Compare-friendly text: 空白差异与多余的成对外层引号不算数据差异。

    Excel和WPS对同一单元格的纯文本可能只差首尾空格、不间断空格或换行符
    （例如“100”对“100 ”）；WPS导出以引号开头的文本时也可能把外层引号吃掉。
    这些都不应让整批上传失败。"""
    text = str(value).translate(_SPACES).replace('\r\n', '\n').replace('\r', '\n')
    text = '\n'.join(_WS_EDGE.sub('', part) for part in text.split('\n')).strip('\n')
    if len(text) >= 2 and text.count('"') == 2 and text.startswith('"') and text.endswith('"'):
        text = text[1:-1]
    return text


# WPS/Excel 公式出错时的显示值：不是真实数据，应视为“无内容”。
_FORMULA_ERRORS = frozenset({
    '#REF!', '#VALUE!', '#DIV/0!', '#N/A', '#NAME?', '#NUM!', '#NULL!',
    '#SPILL!', '#CALC!', '#GETTING_DATA', '#FIELD!', '#BLOCKED!', '#ERROR!',
})

# 常见“无数据”占位符：空单元格、斜杠。它们之间不算内容差异。
_EMPTY_PLACEHOLDERS = frozenset({'', '/'})

# 零时刻（00:00、00:00:00 等）视为“无内容”：表A 用 0/00:00:00 表示“空”，
# 表C 用空白表示“空”，二者等价，不应因此把表A 的 0 清成空白。
_ZERO_TIME = re.compile(r'0+(?::0+)+(?:\.0+)?')

# Excel 序号日期只认 5 位数区间（约 1927~2173 年），避免把 54、11 这类数量误当日期。
_SERIAL_MIN, _SERIAL_MAX = 10000, 200000


def _is_zero(value):
    """数值零（0、0.0、0%…）或零时刻（00:00:00…）都算“无内容”。"""
    number = _numeric_result(value)
    if number is not None:
        return number == 0
    return _ZERO_TIME.fullmatch(normalize_text(value)) is not None


def _no_content(value):
    """空白、公式错误、常见占位符、数值零/零时刻都视为“无内容”，互相之间不算差异。"""
    if normalize_text(value) in _FORMULA_ERRORS | _EMPTY_PLACEHOLDERS:
        return True
    return _is_zero(value)


def _serial_to_date(serial):
    """把 Excel/WPS 序号日期（如 46305）转成 date；越界或非法返回 None。"""
    try:
        return date(1899, 12, 30) + timedelta(days=serial)
    except (ValueError, OverflowError):
        return None


def _parse_date(value):
    """把常见日期写法解析成 (年, 月, 日)；年可为 None。无法识别时返回 None。

    表C 常用 2026/8/6 或 2026-8-6，表A 可能写成 8月6日、2026年8月6日，还可能是
    序号日期 46305（对应 2026/10/10）；这些都是同一日期，只是写法不同。
    """
    if not isinstance(value, str):
        return None
    text = value.strip()
    # 2026/8/6 、2026-8-6 、2026.8.6
    m = re.fullmatch(r'(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})', text)
    if m:
        return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    # 8/6/2026 、08-06-2026 等“月-日-年”写法（WPS 默认 mm-dd-yyyy 的显示）。
    # 表C 用 yyyy/m/d、表A 用 mm-dd-yyyy，都是同一个日期，只是显示格式不同。
    m = re.fullmatch(r'(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})', text)
    if m:
        month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return (year, month, day)
    # 8月6日 、2026年8月6日
    m = re.fullmatch(r'(?:(\d{4})年)?(\d{1,2})月(\d{1,2})日', text)
    if m:
        year = int(m.group(1)) if m.group(1) else None
        return (year, int(m.group(2)), int(m.group(3)))
    # 序号日期 46305（纯数字，且在合理年份区间内）
    if re.fullmatch(r'\d+(?:\.0+)?', text):
        serial = int(Decimal(text))
        if _SERIAL_MIN <= serial <= _SERIAL_MAX:
            d = _serial_to_date(serial)
            if d is not None:
                return (d.year, d.month, d.day)
    return None


def _dates_equivalent(a, b):
    """两个值都是日期、且表示同一天时返回 True（忽略年月日的格式差异）。

    仅当两值都能解析成日期时才比较，避免把普通文本误判为相同；
    任一侧没有年份（如 8月6日）时只比较月与日。
    """
    da = _parse_date(a)
    db = _parse_date(b)
    if da is None or db is None:
        return False
    if da[0] is None or db[0] is None:
        return da[1] == db[1] and da[2] == db[2]
    return da == db


_NUMBER = re.compile(r'^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$')


def _numeric_result(value):
    if isinstance(value, Decimal):
        return value if value.is_finite() else None
    text = normalize_text(value)
    if not _NUMBER.fullmatch(text):
        return None
    # 没有数值类型标记的前导零文本可能是编号，不能把 001 与 1 合并。
    if re.match(r'^[+-]?0\d', text):
        return None
    try:
        number = Decimal(text.rstrip('%').replace(',', ''))
        if text.endswith('%'):
            sign, digits, exponent = number.as_tuple()
            number = Decimal((sign, digits, exponent - 2))
        return number if number.is_finite() else None
    except InvalidOperation:
        return None


def _round_one_decimal(number):
    """只调整比较精度，不修改表格；采用四舍五入，避免浮点数/银行家舍入。"""
    if number.as_tuple().exponent >= -1:
        return number
    with localcontext() as context:
        context.prec = max(28, len(number.as_tuple().digits) + 1)
        return number.quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)


# 表C（如 1 位小数）与表A（如 2 位小数）保存同一个数时，二次舍入会造成最多 0.05 的差
# （例：真实值 158.7487 → 表A 存 158.75，表C 存 158.7）。这种“仅小数位数不同”不是内容
# 差异，直接视为相同、跳过粘贴，不对单元格做任何操作。
_NUMERIC_TOLERANCE = Decimal('0.05')


def values_equal(a, b):
    """公式只比较计算结果；数值四舍五入到小数点后1位再比较，忽略 0.05 内的小数位数差异。"""
    a, b = comparison_result(a), comparison_result(b)
    an, bn = normalize_text(a), normalize_text(b)
    if an == bn or (_no_content(a) and _no_content(b)):
        return True
    av, bv = _numeric_result(a), _numeric_result(b)
    if av is not None and bv is not None:
        if av == bv or _round_one_decimal(av) == _round_one_decimal(bv):
            return True
        # 仅小数位数不同的两个数（差值不超过半个一位小数单位）不算差异。
        return abs(av - bv) <= _NUMERIC_TOLERANCE
    return _dates_equivalent(an, bn)


def differences(expected, actual, origin, limit=None):
    """List the cells that differ between the source batch and what table C holds.

    limit keeps the scan cheap when only the first difference is needed; without
    a limit the count tells the upload loop whether WPS is still filling the
    batch in (fewer differing cells) or the paste has really gone wrong.
    Blank cells, blank rows and whitespace-only differences are normal data."""
    if (expected['row_count'], expected['width']) != (actual['row_count'], actual['width']):
        return [dict(cell='', kind='shape',
                     expected=f"{expected['row_count']:,}行×{expected['width']:,}列",
                     actual=f"{actual['row_count']:,}行×{actual['width']:,}列")]
    col, row = address(origin)
    found = []
    for y, (wanted, got) in enumerate(zip(expected['rows'], actual['rows'])):
        for x in range(expected['width']):
            a = wanted[x] if x < len(wanted) else ''
            b = got[x] if x < len(got) else ''
            if values_equal(a, b):
                continue
            found.append(dict(cell=f'{column_name(col+x)}{row+y}', kind='value',
                              expected=a, actual=b))
            if limit is not None and len(found) >= limit:
                return found
    return found


def mismatch(expected, actual, origin):
    found = differences(expected, actual, origin, limit=1)
    if not found:
        return None
    item = found[0]
    if item['kind'] == 'shape':
        return f"上传后的行列数与本批内容不一致（期望{item['expected']}，实际{item['actual']}）"
    want = str(item['expected'])[:40]
    got = str(item['actual'])[:40]
    return f"表C单元格{item['cell']}与本批内容不一致（期望“{want}”，实际“{got}”）"


def merge_ref(merge, origin):
    col, row = address(origin)
    left, top, right, bottom = merge
    return f'{column_name(col+left)}{row+top}:{column_name(col+right)}{row+bottom}'


def expected_merges(reference, row, count, columns):
    return {(c1-1, r1-row, c2-1, r2-row) for c1,r1,c2,r2 in reference.get('merges', ())
            if 1 <= c1 <= c2 <= columns and row <= r1 <= r2 < row+count}
