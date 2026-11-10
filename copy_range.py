"""Describe which source columns are missing, without guessing or padding data."""
from update_excel import column


def column_name(number):
    result = ''
    while number:
        number, rest = divmod(number - 1, 26)
        result = chr(65 + rest) + result
    return result


def check_filter_column(target, first_col, width, first_row, rows, *, label='剪贴板实际识别范围', mode='selection', requested=None):
    if target.get('mode') != 'filter': return
    name = target['filter_column']
    required = column(name)
    if first_col <= required < first_col + width: return
    last_col, last_row = first_col + width - 1, first_row + rows - 1
    actual = f'{column_name(first_col)}{first_row}:{column_name(last_col)}{last_row}'
    suggestion = f'A1:{column_name(max(last_col,required))}{last_row}'
    details = (f'筛选列为 {name}（第{required}列），{label}为 {actual}'
               f'（{rows:,}行 × {width:,}列），未包含该筛选列。')
    if requested:
        details += f'请求的区域是 {requested}。'
    details += f'请检查筛选列是否正确；若正确，选择“指定区域”，范围至少覆盖 {suggestion}，末行按实际数据填写。'
    if mode == 'all': details += '本次整表操作未取得包含筛选列的完整数据，请确认工作表及网页选区。'
    raise ValueError(details + '表A1未修改。')
