"""Bounded clipboard batches, retaining only one range at a time."""
from update_excel import address
from copy_range import column_name
from excel_clipboard import inspect_copy, CopyIntegrityError
from read_limits import check_cells
from omitted_columns import check_omitted_target, omitted_in_range

DEFAULT_BATCH_ROWS = 1000
MAX_BATCH_CELLS = 200_000


def _batch_rows(config, width):
    """每批行数：默认1000行（可配置1~2000），并按每批最大格数(20万)封顶。"""
    requested = config.get('source', {}).get('copy_batch_rows') or DEFAULT_BATCH_ROWS
    try:
        requested = int(requested)
    except (TypeError, ValueError):
        requested = DEFAULT_BATCH_ROWS
    requested = max(1, min(requested, 2000))
    return min(requested, max(1, MAX_BATCH_CELLS // max(1, width)))


def bounds(ref, config):
    left,right=ref.split(':')
    c1,r1=address(left);c2,r2=address(right)
    if r1!=1 or r2<2 or c2<c1:
        raise ValueError('完整复制范围必须从第1行开始，并包含表头和数据')
    check_cells(r2-r1+1,c2-c1+1,config['source']['max_cells'],'指定区域')
    check_omitted_target(config)
    if len(omitted_in_range(config,c1,c2-c1+1)) == c2-c1+1:
        raise ValueError('指定区域的全部列都已设为舍弃，请保留至少一列数据')
    return c1,r1,c2,r2


def free_bounds(ref, config):
    """任意矩形区域（起点不必是第1行，也不必包含表头）。"""
    left,right=ref.split(':')
    c1,r1=address(left);c2,r2=address(right)
    if r2<r1 or c2<c1:
        raise ValueError('区域终点必须位于起点右下方')
    rows=r2-r1+1;width=c2-c1+1
    check_cells(rows,width,config['source']['max_cells'],'指定区域')
    check_omitted_target(config)
    if len(omitted_in_range(config,c1,width)) == width:
        raise ValueError('指定区域的全部列都已设为舍弃，请保留至少一列数据')
    return c1,r1,c2,r2


def verified_batches(read_range, config, ref, progress, bounds_fn=bounds, retain_rows=False, on_progress=None):
    c1,r1,c2,r2=bounds_fn(ref,config)
    if config.get('source', {}).get('one_shot'):
        # 一次性整块复制：不再分批；若网页复制被截断，自动缩小批次重试仍生效。
        rows_per_batch=r2-r1+1
    else:
        rows_per_batch=_batch_rows(config,c2-c1+1)
    cursor=r1
    while cursor<=r2:
        count=min(rows_per_batch,r2-cursor+1)
        while True:
            first=f'{column_name(c1)}{cursor}'
            part=f'{first}:{column_name(c2)}{cursor+count-1}'
            progress(f'正在复制 {part}；已完成 {cursor-r1:,}/{r2-r1+1:,} 行。全部完成后才保存表A1，请勿操作专用Edge或剪贴板。')
            clip=read_range(part)
            try:
                grid=inspect_copy(clip,config,first,'selection',retain_rows=retain_rows,expected_range=part,batch=True)
            except CopyIntegrityError as exc:
                if exc.can_split and count>1:
                    count=max(1,count//2)
                    rows_per_batch=count
                    progress(f'{part} 复制不完整，正在自动缩小到每批 {count} 行重试；表A1未修改。')
                    continue
                raise ValueError(f'分批复制在 {part} 停止：{exc}') from exc
            yield clip,grid
            if grid.get('synthetic'):
                # 整批都是空白（网页没有可复制的表格）：按请求区域补空白即可，
                # 同时把批次放大，快速跳过末尾的空白行，不必逐行粘贴。
                rows_per_batch=min(max(rows_per_batch*4, 200), _batch_rows(config,c2-c1+1))
                progress(f'{part} 整批为空白单元格，已按空白写入并继续；下一批 {rows_per_batch} 行。')
            cursor+=count
            if on_progress is not None:
                on_progress((cursor-r1)/(r2-r1+1))
            break
