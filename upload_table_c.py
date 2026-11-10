"""把本地表A1按批上传到线上表C的辅助逻辑：分批规划与文本比对。

旧版辅助接口保留兼容。当前上传由upload_runner执行：表C提供格式，
upload_layout规划不切开合并区域的批次，online_format单独设置并核对尺寸。
普通剪贴板粘贴不能保证网页采用列宽和行高。
"""
from copy_range import column_name


def plan_batches(source_start, rows, columns, batch_rows):
    """把 source 范围按行切成批次，返回 (ref, first_row, count) 列表。

    source_start 为 (first_col, first_row)（1基）；每批最多 batch_rows 行。"""
    first_col, first_row = source_start
    batches = []
    for row in range(first_row, first_row + rows, batch_rows):
        count = min(batch_rows, first_row + rows - row)
        ref = f'{column_name(first_col)}{row}:{column_name(first_col + columns - 1)}{row + count - 1}'
        batches.append((ref, row, count))
    return batches


def value_text(value):
    """单元格值转显示文本，用于粘贴后与表A1比对（数字不强制小数位）。"""
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'true' if value else 'false'
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)
