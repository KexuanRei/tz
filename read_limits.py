"""Shared cell-count limits and actionable errors for copy and local reads."""
DEFAULT_CELLS = 100_000_000
MAX_CELLS = 1_000_000_000


class CellLimitError(ValueError):
    def __init__(self, rows, columns, limit, source='复制范围', exact=True):
        cells = rows * columns
        suggested = cells if cells <= MAX_CELLS else None
        self.details = dict(rows=rows, columns=columns, cells=cells,
                            limit=limit, suggested_limit=suggested, exact=exact)
        amount = '' if exact else '至少'
        advice = (f'请将“最大读取单元格数”提高到 {suggested:,} 或以上并保存，再重试。'
                  if suggested else f'已超过支持上限 {MAX_CELLS:,} 格，请指定包含表头的实际数据区域，避免选中整列或空白尾部。')
        super().__init__(f'{source}{amount} {rows:,} 行 × {columns:,} 列 = {cells:,} 格，'
                         f'超过当前读取上限 {limit:,} 格（空白格也计入）。{advice}')


def check_cells(rows, columns, limit, source='复制范围'):
    if rows * columns > min(limit, MAX_CELLS):
        raise CellLimitError(rows, columns, min(limit, MAX_CELLS), source)
