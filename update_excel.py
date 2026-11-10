"""按唯一编号局部更新xlsx，仅修改目标worksheet XML和重算标志。

其他ZIP条目的内容逐字节保留。不会重建整个工作簿的数据模型。
"""
import hashlib
import io
import os
from pathlib import Path
import posixpath
import re
import tempfile
import zipfile
import datetime as dt
from decimal import Decimal
from lxml import etree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def tag(name): return '{' + NS + '}' + name


def column(s):
    if not isinstance(s, str) or not re.fullmatch('[A-Z]{1,3}', s):
        raise ValueError('表B列位置须为A、B、AA等大写字母')
    n = 0
    for ch in s: n = n * 26 + ord(ch) - 64
    if n > 16384: raise ValueError('表B列位置超出Excel范围')
    return n


def address(s):
    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]*)', s or '')
    if not match or int(match[2]) > 1048576:
        raise ValueError('起始单元格格式错误，例如A2')
    return column(match[1]), int(match[2])


def validate_target(t):
    if isinstance(t, dict) and t.get('mode') == 'filter':
        if set(t) != {'mode', 'sheet_name', 'start_cell', 'filter_column', 'filter_value', 'source_start_row'}:
            raise ValueError('筛选复制配置字段不正确')
        if not isinstance(t['sheet_name'], str) or not t['sheet_name'].strip():
            raise ValueError('请填写表B工作表名称')
        address(t['start_cell'])
        column(t['filter_column'])
        if not isinstance(t['filter_value'], str) or not t['filter_value'] or len(t['filter_value']) > 32767:
            raise ValueError('请填写筛选值，例如在建')
        if type(t['source_start_row']) is not int or not 1 <= t['source_start_row'] <= 1048576:
            raise ValueError('表A数据起始行必须是1至1048576的整数')
        return
    if isinstance(t, dict) and t.get('mode') == 'cells':
        if set(t) != {'mode', 'sheet_name', 'cells', 'write_empty'}:
            raise ValueError('单元格复制配置字段不正确')
        if not isinstance(t['sheet_name'], str) or not t['sheet_name'].strip():
            raise ValueError('请填写表B工作表名称')
        if not isinstance(t['write_empty'], bool):
            raise ValueError('空值更新选项无效')
        if not isinstance(t['cells'], list) or not 1 <= len(t['cells']) <= 500:
            raise ValueError('请配置1至500组单元格对应关系')
        destinations = set()
        for item in t['cells']:
            if not isinstance(item, dict) or set(item) != {'row', 'column', 'target'}:
                raise ValueError('每组需填写表A行号、列号及表B单元格')
            for key, limit in [('row', 1048576), ('column', 16384)]:
                if type(item[key]) is not int or not 1 <= item[key] <= limit:
                    raise ValueError('表A行列必须是从1开始的整数，且在Excel范围内')
            address(item['target'])
            if item['target'] in destinations:
                raise ValueError('表B目标单元格不能重复')
            destinations.add(item['target'])
        return
    if not isinstance(t, dict) or set(t) != {'sheet_name', 'start_cell', 'key_source', 'mapping', 'write_empty'}:
        raise ValueError('表B更新配置字段不完整')
    if not isinstance(t['sheet_name'], str) or not t['sheet_name'].strip():
        raise ValueError('请填写表B的实际工作表名称')
    first_col, _ = address(t['start_cell'])
    if not isinstance(t['key_source'], str) or not t['key_source'].strip():
        raise ValueError('请填写唯一编号在表A中的列名')
    mapping = t['mapping']
    if not isinstance(mapping, dict) or not mapping or any(not isinstance(k, str) or not k.strip() for k in mapping):
        raise ValueError('请填写列对应关系，例如 {"订单号":"A","状态":"D"}')
    if t['key_source'] not in mapping:
        raise ValueError('列对应关系必须包含唯一编号列')
    numbers = [column(c) for c in mapping.values()]
    if len(numbers) != len(set(numbers)) or min(numbers) < first_col:
        raise ValueError('目标列不能重复，也不能位于起始单元格左侧')
    if not isinstance(t['write_empty'], bool): raise ValueError('空值更新选项无效')


def xml(raw):
    return ET.fromstring(raw, parser=ET.XMLParser(resolve_entities=False, no_network=True))


def text_value(cell, strings):
    if cell is None: return None
    if cell.find(tag('f')) is not None: raise ValueError('表B唯一编号列不能使用公式')
    kind = cell.get('t')
    v = cell.find(tag('v'))
    value = v.text if v is not None else None
    if kind == 's': return strings[int(value)] if value is not None else None
    if kind == 'inlineStr':
        inline = cell.find(tag('is'))
        return rich_text(inline) if inline is not None else ''
    if kind in ('str', 'd'): return value
    if kind == 'e': raise ValueError('表B唯一编号列含Excel错误值')
    if kind == 'b': return value == '1'
    return Decimal(value) if value is not None else None


def rich_text(node):
    return ''.join(t.text or '' for t in node.xpath('./s:t | ./s:r/s:t', namespaces={'s':NS}))


def key(value):
    if value is None or value == '': return None
    if isinstance(value, bool): return 'true' if value else 'false'
    if isinstance(value, (int, float, Decimal)):
        return format(Decimal(str(value)).normalize(), 'f')
    return str(value)  # 保留前导零及空格，不猜测编号含义


def cell_equal(cell, value, strings):
    old = text_value(cell, strings)
    if old is None and value in (None, ''): return True
    if isinstance(value, bool): return isinstance(old, bool) and old == value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return isinstance(old, Decimal) and old == Decimal(str(value))
    return isinstance(old, str) and isinstance(value, str) and old == value


def update_existing(path, columns, records, target):
    validate_target(target)
    path = Path(path)
    if path.suffix.lower() != '.xlsx' or not path.is_file():
        raise ValueError('请填写已经存在的.xlsx表B完整路径，程序不会自动创建或替换模板')
    if target.get('mode') == 'filter':
        first_col, first_row = address(target['start_cell'])
        rows = list(records)
        width = len(columns)
        if not 1 <= width <= 16384 or first_col + width - 1 > 16384 or first_row + len(rows) - 1 > 1048576:
            raise ValueError('筛选结果超出表B可写入范围，未修改文件')
        if any(len(row) != width for row in rows):
            raise ValueError('筛选行宽度不完整，未修改文件')
        def letter(n):
            result = ''
            while n:
                n, r = divmod(n - 1, 26)
                result = chr(65 + r) + result
            return result
        values = {letter(first_col + j) + str(first_row + i): value
                  for i, row in enumerate(rows) for j, value in enumerate(row)}
        # 本次结果区域完整复制，包括空值；区域以外的历史结果不清空。
        return _update_workbook(path, columns, dict(target, write_empty=True), {}, {}, 1, '', values)
    direct = target.get('mode') == 'cells'
    if direct:
        values = dict(records)
        if set(values) != {item['target'] for item in target['cells']}:
            raise ValueError('单元格读取结果不完整，未修改表B')
        incoming = {}
        mapping = {}
    else:
        return _update_keyed(path, columns, records, target)
    return _update_workbook(path, columns, target, incoming, mapping, 1, '', values)


def _update_keyed(path, columns, records, target):
    mapping = target['mapping']
    if any(h not in columns for h in mapping):
        raise ValueError('列对应关系中的表A列名必须包含在查询导出列中')
    first_col, first_row = address(target['start_cell'])
    key_header = target['key_source']
    source_key_index = columns.index(key_header)
    incoming = {}
    for row in records:
        k = key(row[source_key_index])
        if k is None: raise ValueError('源数据唯一编号为空，未更新表B')
        if k in incoming: raise ValueError('源数据唯一编号重复，未更新表B')
        incoming[k] = row
    return _update_workbook(path, columns, target, incoming, mapping, first_row, key_header)


def _update_workbook(path, columns, target, incoming, mapping, first_row, key_header, values=None, expected_original=None):
    original = path.read_bytes()
    if expected_original is not None and original != expected_original:
        raise ValueError('读取期间目标文件发生修改，已取消写入')
    digest = hashlib.sha256(original).digest()
    archive = zipfile.ZipFile(io.BytesIO(original))
    names = archive.namelist()
    if len(names) != len(set(names)): raise ValueError('xlsx压缩包存在重复条目，停止更新')
    if any(n.startswith('_xmlsignatures/') for n in names):
        raise ValueError('表B带数字签名，更新会使签名失效，请使用未签名副本')
    wb = xml(archive.read('xl/workbook.xml'))
    sheets = wb.find(tag('sheets'))
    sheet = next((s for s in sheets if s.get('name') == target['sheet_name']), None)
    if sheet is None: raise ValueError('表B中找不到指定工作表，请填写底部标签名称')
    rels = xml(archive.read('xl/_rels/workbook.xml.rels'))
    relation = next((r for r in rels if r.get('Id') == sheet.get('{' + REL + '}id')), None)
    if relation is None or relation.get('TargetMode') == 'External': raise ValueError('表B工作表关系无效')
    link = relation.get('Target')
    name = posixpath.normpath(link.lstrip('/') if link.startswith('/') else posixpath.join('xl', link))
    if not name.startswith('xl/worksheets/') or name not in names: raise ValueError('表B工作表路径无效')
    tree = xml(archive.read(name))
    protection = tree.find(tag('sheetProtection'))
    if protection is not None and protection.get('sheet', '1') not in ('0', 'false'):
        raise ValueError('表B工作表已保护，请先在Excel中确认可编辑区域或使用副本')
    strings = []
    if 'xl/sharedStrings.xml' in names:
        strings = [rich_text(si) for si in xml(archive.read('xl/sharedStrings.xml'))]
    data = tree.find(tag('sheetData'))
    if data is None: raise ValueError('目标工作表没有数据区域')
    row_elements = {int(r.get('r')): r for r in data}
    cells = {c.get('r'): c for r in data for c in r if c.tag == tag('c')}
    protected_ranges = []
    merges = tree.find(tag('mergeCells'))
    if merges is not None: protected_ranges += [m.get('ref') for m in merges]
    protected_ranges += [f.get('ref') for f in tree.iter(tag('f')) if f.get('ref')]
    def bounds(ref):
        parts = ref.replace('$', '').split(':')
        a, b = address(parts[0]), address(parts[-1])
        return a[0], a[1], b[0], b[1]
    boxes = [bounds(ref) for ref in protected_ranges]
    def blocked(col, row): return any(a <= col <= c and b <= row <= d for a,b,c,d in boxes)
    dest_keys = {}
    key_col = mapping.get(key_header)
    for row_no in (sorted(row_elements) if values is None else []):
        if row_no < first_row: continue
        if blocked(column(key_col), row_no): raise ValueError('表B编号列与合并或数组公式区域相交')
        k = key(text_value(cells.get(key_col + str(row_no)), strings))
        if k is None: continue
        if k in dest_keys: raise ValueError('表B的唯一编号重复，无法确定要更新哪一行')
        dest_keys[k] = row_no
    tasks = []
    for k, source_row in incoming.items():
        if k not in dest_keys: continue  # 不自动新增行
        row_no = dest_keys[k]
        for header, col in mapping.items():
            if header == key_header: continue  # 编号仅用于匹配，保留原类型/格式
            value = source_row[columns.index(header)]
            tasks.append((col + str(row_no), value))
    if values is not None:
        tasks = list(values.items())
    changed, changed_rows, skipped_formulas = 0, set(), 0
    for ref, value in tasks:
            col_no, row_no = address(ref)
            col = re.match('[A-Z]+', ref)[0]
            if value in (None, '') and not target['write_empty']: continue
            cell = cells.get(ref)
            if blocked(column(col), row_no) or (cell is not None and cell.find(tag('f')) is not None):
                skipped_formulas += 1
                continue
            if cell_equal(cell, value, strings): continue
            row_node = row_elements.get(row_no)
            if row_node is None:
                row_node = ET.Element(tag('row'), r=str(row_no))
                before = next((r for r in data if int(r.get('r')) > row_no), None)
                if before is None: data.append(row_node)
                else: data.insert(data.index(before), row_node)
                row_elements[row_no] = row_node
            if cell is None:
                cell = ET.Element(tag('c'), r=ref)
                before = next((c for c in row_node if c.tag == tag('c') and address(c.get('r'))[0] > column(col)), None)
                if before is not None: row_node.insert(row_node.index(before), cell)
                else: row_node.append(cell)
                cells[ref] = cell
            for child in list(cell):
                if child.tag in (tag('v'), tag('is')): cell.remove(child)
            cell.attrib.pop('t', None)
            if value in (None, ''): pass
            elif isinstance(value, bool):
                cell.set('t', 'b'); ET.SubElement(cell, tag('v')).text = '1' if value else '0'
            elif isinstance(value, (int, float)):
                ET.SubElement(cell, tag('v')).text = str(value)
            else:
                cell.set('t', 'inlineStr')
                inline = ET.SubElement(cell, tag('is'))
                text = ET.SubElement(inline, tag('t'))
                text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                text.text = str(value)
            extension = cell.find(tag('extLst'))
            if extension is not None:
                cell.remove(extension); cell.append(extension)
            changed += 1; changed_rows.add(row_no)
    summary = dict(matched_rows=len(set(incoming) & set(dest_keys)), updated_rows=len(changed_rows),
                   updated_cells=changed, unmatched_source_rows=len(set(incoming)-set(dest_keys)),
                   protected_cells_skipped=skipped_formulas, backup=None)
    if not changed:
        archive.close()
        return summary
    # 现有格式和其他部件保留，公式由Excel下次打开时重算。
    calc = wb.find(tag('calcPr'))
    if calc is None:
        # workbook中的calcPr需位于extLst等尾部元素之前。
        calc = ET.Element(tag('calcPr'))
        following = next((x for x in wb if ET.QName(x).localname in {'oleSize','customWorkbookViews','pivotCaches','smartTagPr','smartTagTypes','webPublishing','fileRecoveryPr','webPublishObjects','extLst'}), None)
        if following is None: wb.append(calc)
        else: wb.insert(wb.index(following), calc)
    calc.set('fullCalcOnLoad', '1'); calc.set('forceFullCalc', '1')
    dimension = tree.find(tag('dimension'))
    if dimension is not None and cells:
        coords = [address(ref) for ref in cells]
        if dimension.get('ref'):
            coords.extend(address(ref) for ref in dimension.get('ref').split(':'))
        def letter(n):
            result = ''
            while n: n, r = divmod(n-1, 26); result = chr(65+r)+result
            return result
        dimension.set('ref', f'{letter(min(c for c,r in coords))}{min(r for c,r in coords)}:{letter(max(c for c,r in coords))}{max(r for c,r in coords)}')
    replacements = {name: ET.tostring(tree, encoding='UTF-8', xml_declaration=True),
                    'xl/workbook.xml': ET.tostring(wb, encoding='UTF-8', xml_declaration=True)}
    fd, tmp = tempfile.mkstemp(prefix='.wps_update_', suffix='.xlsx', dir=path.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, 'w') as output:
            output.comment = archive.comment
            for info in archive.infolist():
                output.writestr(info, replacements.get(info.filename, archive.read(info.filename)))
        if hashlib.sha256(path.read_bytes()).digest() != digest:
            raise ValueError('更新期间表B被其他程序修改，已取消写入，请重试')
        folder = path.parent / (path.stem + '_备份')
        folder.mkdir(exist_ok=True)
        backup = folder / (path.stem + '_' + dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.xlsx')
        with backup.open('xb') as f: f.write(original)
        if hashlib.sha256(path.read_bytes()).digest() != digest:
            raise ValueError('备份期间表B发生修改，已取消写入')
        os.replace(tmp, path)
        summary['backup'] = str(backup)
    finally:
        archive.close()
        if os.path.exists(tmp): os.unlink(tmp)
    return summary
