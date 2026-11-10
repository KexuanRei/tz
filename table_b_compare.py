"""本地表B（线上表A 的 xlsx 快照）读取，以及与表C 的差异单元格对比。

表B 是表A 通过“导出为Excel”下载的完整快照，行/列与线上表A 一致，主键列同表C。
run_sync 用它在本地找出表A 与表C 的差异单元格，再把差异粘贴回线上表A，
从而避免 WatchDog 那样逐段在线读取表A。
"""
from pathlib import Path

from openpyxl import load_workbook

from source_file import render_value
from sync_engine import parse_pk_column, data_start_row, pk_key
from sync_columns import effective_skip_columns
from upload_verify import values_equal


def table_b_path(root):
    return (Path(root) / 'DATA' / '表B.xlsx').resolve()


def read_table_b(root, config, c1, c2):
    """读取 DATA/表B.xlsx，返回 {主键: (表A行号, [c1..c2 列的渲染值])}。

    主键列与跳过列沿用表C 的配置（表A 与表C 结构一致）；数值/日期/百分比按与
    read_table_data 相同的规则渲染成字符串，保证与表C 的对比口径一致。
    """
    path = table_b_path(root)
    if not path.is_file():
        raise ValueError('表B数据文件不存在：' + str(path) + '；请先下载表B（表A 原始 Excel）')
    _pk_name, pk_number = parse_pk_column(config)
    start_row = data_start_row(config)
    skipped = set(effective_skip_columns(config, c1, c2))
    sheet_name = str(config.get('upload', {}).get('sheet_name') or '').strip()
    # data_only=False 让表A 的公式单元格返回公式表达式本身，与表C 的公式对比口径一致。
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        sheet = workbook[sheet_name] if sheet_name in workbook.sheetnames else workbook[workbook.sheetnames[0]]
        rows = {}
        for row_number, cells in enumerate(
                sheet.iter_rows(min_row=start_row, min_col=c1, max_col=c2, values_only=False),
                start=start_row):
            values = []
            key = None
            for col_abs, cell in zip(range(c1, c2 + 1), cells):
                rendered = '' if col_abs in skipped else render_value(
                    getattr(cell, 'value', None), getattr(cell, 'number_format', 'General'))
                if col_abs == pk_number:
                    key = pk_key(rendered)
                values.append(rendered)
            if key is None:
                continue
            if key in rows:
                raise ValueError(f'表A主键“{key}”重复（第{rows[key][0]}、{row_number}行），无法定位')
            rows[key] = (row_number, values)
    finally:
        workbook.close()
    return rows


def table_b_merges(root, config):
    """读取 DATA/表B.xlsx 的合并单元格，返回绝对坐标 {(左列, 上行, 右列, 下行), ...}。

    表B 是表A 的本地快照，其合并单元格与线上表A 一致；线下对比不再在线读取表A，
    因此粘贴前用表B 的合并信息判断目标区域是否涉及合并单元格。
    直接从工作表 XML 读取合并定义，不加载单元格数据，超大表也足够轻量。
    """
    import posixpath
    import zipfile
    from update_excel import xml
    from format_reference import NS, REL_ID, _bounds
    path = table_b_path(root)
    if not path.is_file():
        raise ValueError('表B数据文件不存在：' + str(path) + '；请先下载表B（表A 原始 Excel）')
    sheet_name = str(config.get('upload', {}).get('sheet_name') or '').strip()
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        workbook = xml(archive.read('xl/workbook.xml'))
        rels = xml(archive.read('xl/_rels/workbook.xml.rels'))
        targets = {r.get('Id'): posixpath.normpath(posixpath.join('xl', r.get('Target'))).lstrip('/')
                   for r in rels if r.get('Id') and r.get('Target')}
        target = None
        for sheet in workbook.find(NS + 'sheets'):
            name = sheet.get('name')
            if name == sheet_name:
                target = targets[sheet.get(REL_ID)]
                break
            if not sheet_name and target is None:
                target = targets[sheet.get(REL_ID)]
        if target is None or target not in names:
            raise ValueError('表B中找不到指定工作表“%s”' % sheet_name)
        tree = xml(archive.read(target))
    return {_bounds(m.get('ref')) for m in tree.iter(NS + 'mergeCell') if m.get('ref')}


def target_map_from_table_b(table_b):
    """由表B 的 {主键: (行号, 值)} 得到 {主键: 表A行号}。"""
    return {key: row_number for key, (row_number, _values) in table_b.items()}


def diff_cells(c_values_by_pk, table_b, c1, c2, skipped=()):
    """对比表C 期望值与表B 当前值，返回需要写回表A 的差异单元格。

    c_values_by_pk: {主键: 表C 期望的整行值列表（c1..c2 列，与表B 同宽）}。
    table_b: read_table_b 的返回值 {主键: (表A行号, 表A当前值列表)}。

    返回 [{pk, target_row, cells:[(绝对列号, 表C期望值), ...]}]；
    只列出表C 与表B 都存在的键、且确有单元格差异的行。
    """
    skipped = set(skipped or ())
    result = []
    for key, (row_number, b_values) in table_b.items():
        c_values = c_values_by_pk.get(key)
        if c_values is None:
            continue  # 表B 有而表C 没有：属删除/新增范畴，不在差异单元格范围。
        cells = []
        for i in range(c2 - c1 + 1):
            col_abs = c1 + i
            if col_abs in skipped:
                continue
            c_value = c_values[i] if i < len(c_values) else ''
            b_value = b_values[i] if i < len(b_values) else ''
            if not values_equal(c_value, b_value):
                cells.append((col_abs, c_value))
        if cells:
            result.append(dict(pk=key, target_row=row_number, cells=cells))
    return result
