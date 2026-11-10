"""新增记录插入的纯逻辑：找最近可映射主键锚点，按 Δ行 定位表A目标行。

本模块只服务“新增”：表C2 出现表A2 没有的主键时，以上下最近的已有主键为锚点，
目标行 = 表A2 锚点行 + (本行 - 锚点行)；随后把新记录的非空单元格（含新主键）写入该空白行。

“修改”（已有主键的记录变化）仍沿用 sync_runner 的整行复制，不经过本模块。
"""
from upload_verify import normalize_text, comparison_result, values_equal


def find_inserts(pk_rows, a_index):
    """找出需要插入的新记录（主键不在 a_index），并确定目标行。

    pk_rows: 按行号升序的 [(sheet_row, 主键或None), ...]。
    a_index: {主键值: 表A2行号}。

    返回 list of dict：source_row, pk, target_row, anchor_pk, anchor_source_row, delta_row。
    找不到锚点的项 target_row / anchor_* 为 None（由调用方统计为“无法定位”）。
    """
    pks = [pk for _row, pk in pk_rows]
    out = []
    for i, (row, pk) in enumerate(pk_rows):
        if pk is None or pk in a_index:
            continue
        anchor = _nearest_mappable(pks, pk_rows, i, a_index)
        if anchor is None:
            out.append(dict(source_row=row, pk=pk, target_row=None,
                            anchor_pk=None, anchor_source_row=None, delta_row=None))
            continue
        _j, anchor_row, anchor_pk = anchor
        out.append(dict(source_row=row, pk=pk,
                        target_row=a_index[anchor_pk] + (row - anchor_row),
                        anchor_pk=anchor_pk, anchor_source_row=anchor_row,
                        delta_row=row - anchor_row))
    return out


def _nearest_mappable(pks, pk_rows, index, a_index):
    """返回 (j, anchor_row, anchor_pk)：离 index 最近、且主键值在 a_index 中的行；
    上下各取最近者，取行距较小的一方（相等取上方）。找不到返回 None。"""
    up = down = None
    for j in range(index - 1, -1, -1):
        if pks[j] is not None and pks[j] in a_index:
            up = (j, pk_rows[j][0], pks[j])
            break
    for j in range(index + 1, len(pk_rows)):
        if pks[j] is not None and pks[j] in a_index:
            down = (j, pk_rows[j][0], pks[j])
            break
    if up is None:
        return down
    if down is None:
        return up
    row = pk_rows[index][0]
    if (row - up[1]) <= (down[1] - row):
        return up
    return down


def changed_cells(d_values, a_values, first_col, width, skipped):
    """返回应写入表A 的 (绝对列号, 表C值) 列表。

    只做“填入 / 更新”：表C 为空白时跳过（不清空表A 已有内容）；
    跳过列与归一化后无差异的单元格也排除。新记录目标行为空白时，等价于所有非空表C单元格。

    skipped: 绝对 1 基列号的集合（不读取、不比较、不写入）。
    """
    skipped = skipped or frozenset()
    out = []
    for i in range(width):
        col_abs = first_col + i
        if col_abs in skipped:
            continue
        d = d_values[i] if i < len(d_values) else ''
        a = a_values[i] if i < len(a_values) else ''
        if not normalize_text(comparison_result(d)):
            continue
        if values_equal(d, a):
            continue
        out.append((col_abs, d))
    return out
