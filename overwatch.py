"""表C → 持久化表A1镜像 → 按主键写回表A。"""
import json
from pathlib import Path

from openpyxl import load_workbook
from local_mirror import mirror_path, STATE_SHEET


class OverWatcher:
    def __init__(self, copier, config):
        self.copier, self.config = copier, config

    def _key(self):
        return json.dumps({k: self.config.get(k) for k in ('local', 'source', 'sync')}, sort_keys=True)

    def ready(self):
        key = self._key()
        path = mirror_path(getattr(self.copier, 'root', Path(__file__).parent), self.config)
        cached = getattr(self.copier, '_a1_mirror', None)
        if cached is not None and cached.path == path and cached.initialized and not path.exists():
            raise ValueError('表A1镜像文件已被移走或删除，请恢复文件后重试')
        if (getattr(self.copier, '_overwatch_built', False)
                and getattr(self.copier, '_a1_ready_key', key) == key):
            return True
        if not path.exists():
            return False
        book = load_workbook(path, read_only=True)
        try:
            if STATE_SHEET not in book or self.config.get('local', {}).get('sheet_name') not in book:
                # 用户自己指定的文件，尚未被本程序初始化为镜像：需要先构建（会就地收养）。
                return False
            metadata = json.loads(book[STATE_SHEET].cell(1, 1).value or '{}')
            ready = bool(metadata.get('initialized'))
        finally:
            book.close()
        if ready:
            self.copier._overwatch_built = True
            self.copier._a1_ready_key = key
        return ready

    def build(self):
        from sync_runner import build_mirror
        result = build_mirror(self.copier, self.config)
        self.copier._overwatch_built = True
        self.copier._a1_ready_key = self._key()
        return result

    def watch_once(self):
        from sync_runner import run_sync
        return run_sync(self.copier, self.config)

    def run_once(self):
        return self.watch_once() if self.ready() else self.build()
