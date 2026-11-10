"""Raise only the saved cell allowance, preserving the rest of the user's config."""
import datetime as dt
import json
import os
from pathlib import Path
import tempfile

from read_limits import DEFAULT_CELLS, MAX_CELLS


def raise_limit(root, value=DEFAULT_CELLS):
    if type(value) is not int or not 1 <= value <= MAX_CELLS:
        raise ValueError('读取上限必须为1到10亿之间的整数')
    path = Path(root)/'config.json'
    original = path.read_bytes()
    config = json.loads(original.decode('utf-8-sig'))
    source = config.get('source')
    if not isinstance(source, dict): raise ValueError('config.json缺少source配置，未修改')
    current = source.get('max_cells')
    if type(current) is not int or not 1 <= current <= MAX_CELLS:
        raise ValueError('现有读取上限无效，请在网页高级设置中修改')
    if current >= value: return current, None
    source['max_cells'] = value
    updated = json.dumps(config, ensure_ascii=False, indent=2).encode('utf-8')
    fd, temp = tempfile.mkstemp(prefix='.config_limit_', suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'wb') as handle: handle.write(updated)
        if path.read_bytes() != original: raise ValueError('配置正在被其他程序修改，请关闭服务后重试')
        backup = path.with_name('config.读取上限备份.'+dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.json')
        with backup.open('xb') as handle: handle.write(original)
        if path.read_bytes() != original: raise ValueError('配置发生变化，已取消替换')
        os.replace(temp, path)
        return value, backup
    finally:
        if os.path.exists(temp): os.unlink(temp)


def main():
    try:
        value, backup = raise_limit(Path(__file__).resolve().parent)
    except (OSError, ValueError):
        print('调整未完成。请确认程序目录存在有效的config.json，关闭服务后重试；也可在网页高级设置中修改。')
        return 1
    print(f'当前已保存的读取上限：{value:,} 格。')
    if backup: print('原配置已备份为：'+backup.name)
    print('请重新启动start_ui.bat，并在网页按Ctrl+F5刷新后重新复制。')
    return 0


if __name__ == '__main__': raise SystemExit(main())
