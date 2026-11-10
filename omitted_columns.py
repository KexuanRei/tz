"""Explicitly omitted browser columns; preserve original spreadsheet coordinates."""
from update_excel import column
from copy_range import column_name


def parse_omitted_columns(value):
    if not isinstance(value, str):
        raise ValueError('舍弃的隐藏列须填写列字母，例如 F,W,AH；不舍弃时留空')
    if not value.strip():
        return ()
    names = [part.strip().upper() for part in value.replace('，', ',').split(',')]
    try:
        numbers = tuple(sorted(column(name) for name in names))
    except ValueError as exc:
        raise ValueError('舍弃的隐藏列格式错误，请填写 F,W,AH 这样的列字母，用逗号分隔') from exc
    if len(numbers) != len(set(numbers)):
        raise ValueError('舍弃的隐藏列不能重复')
    return numbers


def omitted_in_range(config, first_col, width):
    columns = parse_omitted_columns(config['source'].get('omitted_columns', ''))
    return tuple(c for c in columns if first_col <= c < first_col + width)


def check_omitted_target(config):
    return parse_omitted_columns(config['source'].get('omitted_columns', ''))


def expand_values(values, first_col, width, omitted):
    """Used only when a caller requests retained TSV values, not in large batches."""
    source = iter(values)
    omitted = set(omitted)
    return ['' if c in omitted else next(source, '') for c in range(first_col, first_col + width)]


def column_list(columns):
    return ','.join(column_name(c) for c in columns)
