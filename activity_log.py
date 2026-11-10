"""程序自监控：把所有操作记录到程序根目录 LOG 文件夹下的文本日志。

网页「操作记录」开关（config.monitor.enabled）控制是否写入；关闭时不产生任何日志文件。
每天一个日志文件（LOG/YYYY-MM-DD.log），一行一条事件，敏感信息（网址、密码、令牌等）脱敏。
写日志失败不影响同步本身。
"""
import datetime as dt
import threading
from pathlib import Path

from diagnostics import scrub

# 最多保留最近多少个日志文件，更早的自动删除，避免 LOG 无限膨胀。
MAX_FILES = 90


class ActivityLog:
    def __init__(self, root):
        self.root = Path(root)
        self.folder = self.root / 'LOG'
        self.lock = threading.RLock()
        self.enabled = False

    def configure(self, config):
        """按最新配置更新开关；config 缺省 monitor 时视为关闭。"""
        monitor = config.get('monitor') if isinstance(config, dict) else None
        self.enabled = bool((monitor or {}).get('enabled', False))

    def record(self, action, message='', **fields):
        if not self.enabled:
            return
        now = dt.datetime.now().astimezone()
        parts = [f'[{now.strftime("%Y-%m-%d %H:%M:%S")}] {scrub(str(action))}']
        if message:
            parts.append(scrub(str(message)))
        for key, value in fields.items():
            if value is None:
                continue
            if isinstance(value, (dict, list, tuple)):
                value = _flatten(value)
            parts.append(f'{scrub(str(key))}={scrub(str(value))}')
        line = ' | '.join(parts) + '\n'
        try:
            with self.lock:
                self.folder.mkdir(parents=True, exist_ok=True)
                path = self.folder / (now.strftime('%Y-%m-%d') + '.log')
                with path.open('a', encoding='utf-8') as stream:
                    stream.write(line)
                self._prune()
        except OSError:
            pass

    def _prune(self):
        try:
            logs = sorted(self.folder.glob('*.log'))
            for old in logs[:-MAX_FILES]:
                old.unlink(missing_ok=True)
        except OSError:
            pass


def _flatten(value):
    """把 dict/list 转成紧凑的单行文本，避免日志里出现多行或过长的结构。"""
    try:
        import json
        text = json.dumps(value, ensure_ascii=False, default=str)
        return text if len(text) <= 2000 else text[:2000] + '…'
    except Exception:
        return str(value)[:2000]
