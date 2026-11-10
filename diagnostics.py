"""Small, local error reports; never collect request bodies or workbook contents."""
import datetime as dt
import json
import os
from pathlib import Path
import re
import sys
import threading
import traceback
import uuid

from read_limits import CellLimitError

VERSION = '2026.10.30.merged-headers'
CAPABILITIES = ['sync', 'local_mirror', 'overwatch', 'watchdog', 'diagnostics', 'download_table_c', 'filtered_mirror', 'cell_formats']


def scrub(value):
    text = str(value)
    for key in ('WPS_TOKEN',):
        secret = os.environ.get(key)
        if secret:
            text = text.replace(secret, '[已隐藏]')
    text = re.sub(r'(?i)((?:password|passwd|token|authorization|cookie|密码|令牌)[\s\"\x27]*[:=]\s*)(?:\"[^\"]*\"|\x27[^\x27]*\x27|[^\s,;]+)', r'\1[已隐藏]', text)
    text = re.sub(r'https?://[^\s<>\"\x27]+', '[网页地址已隐藏]', text)
    return text[:12000]


def error_code(exc):
    chain, seen = [], set()
    while exc is not None and id(exc) not in seen:
        seen.add(id(exc)); chain.append(exc)
        exc = exc.__cause__ or (None if exc.__suppress_context__ else exc.__context__)
    for item in chain:
        if getattr(item, 'code', None) in {'TASK_BUSY','SHEET_TAB_NOT_FOUND','SHEET_TAB_AMBIGUOUS','SHEET_TAB_UNCONFIRMED',
            'NAMEBOX_AMBIGUOUS','NAMEBOX_NOT_FOUND','NAMEBOX_FOCUS_LOST','NAMEBOX_NOT_EDITABLE',
            'NAMEBOX_INPUT_REJECTED','NAMEBOX_INTERACTION_FAILED','RANGE_NOT_CONFIRMED',
            'UPLOAD_TARGET_INVALID','UPLOAD_TARGET_TOO_SMALL','UPLOAD_TARGET_UNVERIFIED',
            'DIMENSION_CONTROL_UNAVAILABLE','DIMENSION_PANEL_MISMATCH'}:
            return item.code
        if type(item).__name__ == 'TargetClosedError': return 'BROWSER_CLOSED'
        if isinstance(item, ModuleNotFoundError): return 'DEPENDENCY_MISSING'
        if getattr(item, 'winerror', None) == 10048 or getattr(item, 'errno', None) in (98,10048): return 'PORT_IN_USE'
        if isinstance(item, MemoryError): return 'MEMORY_EXHAUSTED'
        if isinstance(item, CellLimitError): return 'CELL_LIMIT_EXCEEDED'
        if isinstance(item, PermissionError): return 'FILE_ACCESS_DENIED'
        if isinstance(item, FileNotFoundError): return 'FILE_NOT_FOUND'
        if 'Timeout' in type(item).__name__: return 'OPERATION_TIMEOUT'
    return 'VALIDATION_FAILED' if chain and isinstance(chain[0], ValueError) else 'UNEXPECTED_ERROR'


class Reports:
    def __init__(self, root):
        self.root = Path(root)
        self.folder = self.root / 'error_reports'
        self.lock = threading.RLock()
        self.memory = []

    def _duplicate(self, code, stage, message, now):
        """An identical (code, stage, message) failure within half an hour
        reuses its earlier report instead of piling up near-identical files.
        Never matches across different messages or stages."""
        key = (code, stage, message)
        candidates = list(self.memory)
        try:
            for path in sorted(self.folder.glob('*.json'), reverse=True)[:30]:
                try:
                    if path.stat().st_size > 100000:
                        continue
                    candidates.append(json.loads(path.read_text(encoding='utf-8')))
                except (OSError, ValueError):
                    pass
        except OSError:
            pass
        for report in candidates:
            if not isinstance(report, dict):
                continue
            try:
                stamp = dt.datetime.fromisoformat(report['time'])
            except (KeyError, TypeError, ValueError):
                continue
            if (report.get('code'), report.get('stage'), report.get('message')) == key and \
                    0 <= (now - stamp).total_seconds() < 1800:
                return report
        return None

    def record(self, exc, stage, *, message=None, code=None, context=None):
        now = dt.datetime.now(dt.timezone.utc)
        report = dict(id=now.strftime('%Y%m%dT%H%M%S%f') + '-' + uuid.uuid4().hex[:10],
                      time=now.isoformat(), version=VERSION, code=code or error_code(exc),
                      stage=scrub(stage), message=scrub(message if message is not None else
                      (str(exc) if isinstance(exc, ValueError) else type(exc).__name__)),
                      python=sys.version.split()[0], executable=sys.executable,
                      program_directory=str(self.root.resolve()), pid=os.getpid(),
                      context={str(k): scrub(v) for k,v in (context or {}).items()}, exceptions=[])
        # Frame locations and types are useful without source lines, locals, SQL,
        # raw COM/Playwright arguments, request bodies, cookies or clipboard data.
        seen = set()
        while exc is not None and id(exc) not in seen and len(seen) < 10:
            seen.add(id(exc))
            item = dict(type=type(exc).__module__ + '.' + type(exc).__name__,
                        stack=[dict(file=Path(f.filename).name, line=f.lineno, function=f.name)
                               for f in traceback.extract_tb(exc.__traceback__)[-24:]])
            for attr in ('errno', 'winerror', 'hresult'):
                number = getattr(exc, attr, None)
                if isinstance(number, int): item[attr] = number
            if exc.args and isinstance(exc.args[0], int): item['native_code'] = exc.args[0]
            if isinstance(exc, ModuleNotFoundError): item['module'] = scrub(exc.name or '')
            details = getattr(exc, 'diagnostic_context', None)
            if isinstance(details, dict): item['ui_context'] = {str(k):scrub(v) for k,v in details.items()}
            report['exceptions'].append(item)
            exc = exc.__cause__ or (None if exc.__suppress_context__ else exc.__context__)
        with self.lock:
            existing = self._duplicate(report['code'], report['stage'], report['message'], now)
            if existing is not None:
                return existing
            try:
                self.folder.mkdir(exist_ok=True)
                path = self.folder / (report['id'] + '.json')
                temp = path.with_suffix('.tmp')
                temp.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
                os.replace(temp, path)
                for old in sorted(self.folder.glob('*.json'), reverse=True)[30:]:
                    old.unlink(missing_ok=True)
            except OSError:
                report['storage_warning'] = '报告无法写入本地文件，仍可在网页复制或下载。'
            self.memory.append(report)
            self.memory = self.memory[-30:]
        return report

    def recent(self, limit=20):
        with self.lock:
            reports = {r['id']: r for r in self.memory}
            try:
                for path in sorted(self.folder.glob('*.json'), reverse=True)[:30]:
                    try:
                        if path.stat().st_size > 100000: continue
                        report = json.loads(path.read_text(encoding='utf-8'))
                        if isinstance(report, dict) and isinstance(report.get('id'), str):
                            reports[report['id']] = report
                    except (OSError, ValueError): pass
            except OSError: pass
            return sorted(reports.values(), key=lambda r:r.get('time',''), reverse=True)[:limit]

    def latest(self):
        return next(iter(self.recent(1)), None)
