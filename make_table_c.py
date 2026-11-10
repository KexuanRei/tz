"""Build a format-only 表C workbook from a local copy of the source sheet.

表C keeps only row heights, column widths, merged ranges and the frozen pane;
all cell values and formulas are dropped, so it never carries or leaks data.
Usage: 生成表C.bat 表A副本.xlsx 表C.xlsx [工作表名]
"""
import argparse
from pathlib import Path

from copy_range import column_name
from format_reference import load


def create(source, target, sheet_name='', target_sheet='表C'):
    source = Path(source).resolve()
    target = Path(target).resolve()
    if source == target:
        raise ValueError('表C不能覆盖源文件，请另选输出路径')
    if source.suffix.lower() != '.xlsx' or not source.is_file():
        raise ValueError('请提供已有的.xlsx表A副本作为来源')
    if target.suffix.lower() != '.xlsx':
        raise ValueError('表C输出路径必须以.xlsx结尾')
    config = {'source': {'local_file': '', 'format_file': str(source), 'format_sheet': sheet_name}}
    reference = load(Path('.'), config)
    if reference is None:
        raise ValueError('未能读取来源文件的格式')
    from openpyxl import Workbook
    book = Workbook()
    sheet = book.active
    sheet.title = target_sheet or '表C'
    for col, width in reference['column_widths'].items():
        sheet.column_dimensions[column_name(col)].width = width
    for row, height in reference['row_heights'].items():
        sheet.row_dimensions[row].height = height
    for c1, r1, c2, r2 in reference['merges']:
        sheet.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    if reference['pane']:
        sheet.freeze_panes = reference['pane'][2]
    if reference['default_row_height'] is not None:
        sheet.sheet_format.defaultRowHeight = reference['default_row_height']
        sheet.sheet_format.customHeight = True
    if reference['default_col_width'] is not None:
        sheet.sheet_format.defaultColWidth = reference['default_col_width']
    target.parent.mkdir(parents=True, exist_ok=True)
    book.save(target)
    return dict(widths=len(reference['column_widths']), heights=len(reference['row_heights']),
                merges=len(reference['merges']), pane=reference['pane'])


def create_from_layout(layout, target, sheet_name='表C'):
    """用在线页面提取的格式信息生成只含格式的 表C.xlsx。

    layout 为 format_reference.load 或 page_layout.extract 返回的格式字典
    （column_widths/row_heights/merges/pane/default_row_height/default_col_width）。
    只写格式、不写任何单元格数据，避免泄露数据。
    """
    target = Path(target).resolve()
    if target.suffix.lower() != '.xlsx':
        raise ValueError('表C输出路径必须以.xlsx结尾')
    from openpyxl import Workbook
    book = Workbook()
    sheet = book.active
    sheet.title = sheet_name or '表C'
    for col, width in (layout.get('column_widths') or {}).items():
        if width is not None:
            sheet.column_dimensions[column_name(col)].width = width
    for row, height in (layout.get('row_heights') or {}).items():
        if height is not None:
            sheet.row_dimensions[row].height = height
    for merge in (layout.get('merges') or ()):
        c1, r1, c2, r2 = merge
        if c1 != c2 or r1 != r2:
            sheet.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    if layout.get('pane'):
        sheet.freeze_panes = layout['pane'][2]
    if layout.get('default_row_height') is not None:
        sheet.sheet_format.defaultRowHeight = layout['default_row_height']
        sheet.sheet_format.customHeight = True
    if layout.get('default_col_width') is not None:
        sheet.sheet_format.defaultColWidth = layout['default_col_width']
    target.parent.mkdir(parents=True, exist_ok=True)
    book.save(target)
    return dict(widths=len(layout.get('column_widths') or {}),
                heights=len(layout.get('row_heights') or {}),
                merges=len(layout.get('merges') or ()),
                pane=layout.get('pane'))


def main():
    parser = argparse.ArgumentParser(description='从表A副本生成只含格式的表C')
    parser.add_argument('source', help='表A副本.xlsx完整路径')
    parser.add_argument('target', help='表C输出.xlsx完整路径')
    parser.add_argument('sheet', nargs='?', default='', help='来源工作表名（默认第一个工作表）')
    args = parser.parse_args()
    if not args.source or not args.target:
        print('用法：生成表C.bat 表A副本.xlsx 表C.xlsx [工作表名]')
        return 1
    try:
        result = create(args.source, args.target, args.sheet)
    except (ValueError, OSError) as exc:
        print('生成失败：' + str(exc))
        return 1
    print('表C已生成：列宽%s、行高%s、合并区域%s、冻结窗格%s。' % (
        result['widths'], result['heights'], result['merges'],
        result['pane'][2] if result['pane'] else '无'))
    print('在网页“表C路径”中填写该文件；程序只从表C读取格式，绝不读取数据。')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
