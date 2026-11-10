"""Read only layout from a local reference workbook: column widths, row
heights, merged ranges and frozen panes. No cell data is read from this file;
data always comes from the live copy. Coordinates are absolute worksheet
positions, so they apply directly to the pasted range."""
import posixpath
import re
import zipfile
from pathlib import Path

from lxml import etree

from update_excel import xml  # safe parser: no entity resolution, no network


NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL_ID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'


def _bounds(ref):
    match = re.fullmatch(r'([A-Z]{1,3})([0-9]+):([A-Z]{1,3})([0-9]+)', ref or '')
    if not match:
        raise ValueError('格式参考文件包含无效的合并区域')
    def column(s):
        n = 0
        for ch in s:
            n = n * 26 + ord(ch) - 64
        return n
    return column(match[1]), int(match[2]), column(match[3]), int(match[4])


def load(root, config):
    """Return layout info or None when no reference file is configured."""
    value = (config['source'].get('format_file') or '').strip()
    if not value:
        return None
    path = (Path(root) / value).resolve()
    if path.suffix.lower() != '.xlsx' or not path.is_file():
        raise ValueError('表A格式参考文件不存在或不是.xlsx文件，请检查路径；不需要时留空')
    if path == (Path(root) / config['source'].get('local_file', '')).resolve():
        raise ValueError('表A格式参考文件不能是表A1本身，请选择结构相同的其他.xlsx副本')
    sheet_name = (config['source'].get('format_sheet') or '').strip()
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if 'xl/workbook.xml' not in names or 'xl/_rels/workbook.xml.rels' not in names:
            raise ValueError('格式参考文件不是有效的工作簿')
        workbook = xml(archive.read('xl/workbook.xml'))
        rels = xml(archive.read('xl/_rels/workbook.xml.rels'))
        targets = {r.get('Id'): posixpath.normpath(posixpath.join('xl', r.get('Target'))).lstrip('/')
                   for r in rels if r.get('Id') and r.get('Target')}
        target, used_name = None, None
        for sheet in workbook.find(NS + 'sheets'):
            name = sheet.get('name')
            if name == sheet_name:
                target, used_name = targets[sheet.get(REL_ID)], name
            elif not sheet_name and target is None:
                target, used_name = targets[sheet.get(REL_ID)], name
        if target is None or target not in names:
            available = [s.get('name') for s in workbook.find(NS + 'sheets')]
            detail = '，该文件包含的工作表：' + '、'.join(available) if available else '，该文件没有任何工作表'
            raise ValueError('表C中找不到指定工作表“%s”%s，请填写该文件底部标签名称' % (sheet_name, detail))
        tree = xml(archive.read(target))
    column_widths = {}
    cols = tree.find(NS + 'cols')
    if cols is not None:
        for entry in cols.findall(NS + 'col'):
            if entry.get('width') is None:
                continue
            width = float(entry.get('width'))
            for number in range(int(entry.get('min')), int(entry.get('max')) + 1):
                column_widths[number] = width
    row_heights = {int(r.get('r')): float(r.get('ht'))
                   for r in tree.iter(NS + 'row') if r.get('ht') is not None}
    merges = [_bounds(m.get('ref')) for m in tree.iter(NS + 'mergeCell') if m.get('ref')]
    pane = None
    for entry in tree.iter(NS + 'pane'):
        if entry.get('state') == 'frozen' and entry.get('topLeftCell'):
            pane = (int(entry.get('xSplit') or 0), int(entry.get('ySplit') or 0),
                    entry.get('topLeftCell'))
    defaults = tree.find(NS + 'sheetFormatPr')
    default_row_height = default_col_width = None
    if defaults is not None:
        if defaults.get('defaultRowHeight') is not None:
            default_row_height = float(defaults.get('defaultRowHeight'))
        if defaults.get('defaultColWidth') is not None:
            default_col_width = float(defaults.get('defaultColWidth'))
    return dict(sheet=used_name, column_widths=column_widths, row_heights=row_heights,
                merges=merges, pane=pane,
                default_row_height=default_row_height, default_col_width=default_col_width)
