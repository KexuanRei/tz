"""Read the saved local A1 workbook, without modifying it or contacting WPS."""
from pathlib import Path


def source_path(root, config):
    value = config['source'].get('local_file', '')
    if not isinstance(value, str) or not value.strip():
        raise ValueError('请填写本地表A1的完整.xlsx路径')
    path = (Path(root) / value).resolve()
    if path.suffix.lower() != '.xlsx' or not path.is_file():
        raise ValueError('表A1文件不存在或不是.xlsx文件，请粘贴数据并保存后重试')
    return path
