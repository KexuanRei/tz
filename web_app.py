"""配置及运行界面。默认仅本机访问；部署到服务器见 部署到服务器教程.md。

启动：python web_app.py（默认 http://127.0.0.1:8765）
远程/无人值守：python web_app.py --host 0.0.0.0 --allow-host 你的域名 --auth 用户:密码 --auto-start

分批比较在线表C与表A，连续差异行合并复制并校验；整轮结束后等待设置的秒数。
"""
import argparse
import json
import os
from pathlib import Path
import re
import secrets
import sys
import threading
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from read_limits import MAX_CELLS, CellLimitError
from update_excel import column
from upload_target import parse_destination
from diagnostics import Reports, VERSION, CAPABILITIES
from sync_engine import sync_batch_rows, MAX_BATCH_ROWS
from sync_columns import sync_skip_columns, format_skip_columns, parse_sync_columns

ROOT = Path(__file__).resolve().parent

# 旧版（表A→表A1→表C）字段：加载旧配置时迁移为新schema，不再参与任何流程。
LEGACY_MARKERS = ('schedule', 'export', 'index_columns')


def _defaults():
    return {
        "source": {"share_url": "", "sheet_name": "Sheet1", "max_cells": 100_000_000,
                   "scan_range": "", "batch_rows": 1000,
                   "format_file": "", "format_sheet": ""},
        "local": {"local_file": "表A1.xlsx", "sheet_name": "表A1"},
        "filter": {"enabled": False, "sheet_name": "系统配置", "header_row": 1,
                   "local_file": "DATA/筛选镜像.xlsx", "mirror_sheet": "筛选镜像"},
        "upload": {"share_url": "", "sheet_name": "Sheet1", "scan_range": ""},
        "sync": {"enabled": True, "wait_seconds": 120, "max_pastes": 0, "primary_key_column": "D", "data_start_row": 2,
                 "batch_rows": 100, "watch_enabled": False, "watch_debounce": 1.0,
                 "skip_columns": "", "sync_columns": "", "insert_enabled": False,
                 "start_cell": "A5", "legacy_scan": False},
        "monitor": {"enabled": True},
    }


def _migrate_legacy(c):
    """旧配置（表A→表A1→表C）一次性迁移到新schema。

    表C/表A 分享链接角色互换，为避免指向错误的表，两者都留空由用户重新填写；
    保留表A1路径、表A1工作表名与读取上限。
    """
    source = c.get('source', {})
    if not any(k in c for k in LEGACY_MARKERS) and 'local_file' not in source:
        return None
    new = _defaults()
    if source.get('local_file'):
        new['local']['local_file'] = str(source['local_file'])
    if source.get('sheet_name'):
        new['local']['sheet_name'] = str(source['sheet_name'])
    if isinstance(source.get('max_cells'), (int, float)):
        new['source']['max_cells'] = int(source['max_cells'])
    return new


def _share_link(value, label):
    if not value:
        return ''
    share = urllib.parse.urlparse(value)
    if share.scheme != 'https' or share.hostname not in {'www.kdocs.cn', 'kdocs.cn'} \
            or not share.path.startswith('/l/') or share.username or share.password:
        raise ValueError(f'{label}应为https://www.kdocs.cn/l/…形式')
    return value


def _rectangle(value, label):
    value = str(value or '').strip().upper()
    if not value:
        return ''
    if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*', value):
        raise ValueError(f'{label}需填写完整矩形区域，例如 A1:CB50000；留空表示自动整表')
    target = parse_destination(value)
    if not target.bounded:
        raise ValueError(f'{label}需填写完整矩形区域（起止都要有），例如 A1:CB50000')
    return target.ref


def validate_config(c):
    if not isinstance(c, dict):
        raise ValueError('配置必须是JSON对象')
    defaults = _defaults()
    for section, values in defaults.items():
        c.setdefault(section, {})
        if not isinstance(c[section], dict):
            raise ValueError('配置字段不正确：' + section)
        for key, value in values.items():
            c[section].setdefault(key, value)
    # 移除误加的整轮扫描上限；本版本始终扫描完整区域。
    c['sync'].pop('scan_rows', None)
    schema = {key: set(defaults[key]) for key in defaults}
    if set(c) != set(schema):
        raise ValueError('配置字段不完整或包含未知字段')
    for section, keys in schema.items():
        if set(c[section]) != keys:
            raise ValueError('配置字段不正确：' + section)

    def number(section, key, low, high, integer=False):
        value = c[section][key]
        if isinstance(value, bool) or not isinstance(value, (int, float)) \
                or not low <= value <= high or (integer and int(value) != value):
            raise ValueError(key + '超出允许范围')

    number('source', 'max_cells', 1, MAX_CELLS, True)
    number('source', 'batch_rows', 10, MAX_BATCH_ROWS, True)
    # 整轮扫描完成后的等待时间（秒）：其他人的修改只能靠这一轮轮扫描发现。
    number('sync', 'wait_seconds', 1, 86400, True)
    # 每轮最多同步多少个变更行（0=不限，一轮扫完整表）。
    number('sync', 'max_pastes', 0, 1000, True)
    number('sync', 'data_start_row', 1, 1048576, True)
    c['sync']['batch_rows'] = sync_batch_rows(c)

    start_cell_value = str(c['sync'].get('start_cell') or 'A5').strip().upper().replace('$', '')
    if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*', start_cell_value):
        raise ValueError('起始单元格须填写单元格地址，例如 A5')
    c['sync']['start_cell'] = start_cell_value
    if not isinstance(c['sync'].get('legacy_scan'), bool):
        raise ValueError('旧检索方式开关必须为开关值')

    if not isinstance(c['sync']['enabled'], bool):
        raise ValueError('同步开关必须为开关值')
    if not isinstance(c['sync'].get('watch_enabled'), bool):
        raise ValueError('即时同步开关必须为开关值')
    if not isinstance(c['sync'].get('insert_enabled'), bool):
        raise ValueError('新增记录插入开关必须为开关值')
    number('sync', 'watch_debounce', 0.3, 10.0, False)

    if not isinstance(c['monitor'].get('enabled'), bool):
        raise ValueError('操作记录开关必须为开关值')

    for section in ('source', 'local', 'upload'):
        name = c[section]['sheet_name']
        if not isinstance(name, str) or not name.strip():
            raise ValueError('工作表名称不能为空')
        c[section]['sheet_name'] = str(name).strip()

    c['source']['share_url'] = _share_link(c['source'].get('share_url') or '', '表C分享链接')
    c['upload']['share_url'] = _share_link(c['upload'].get('share_url') or '', '表A分享链接')

    c['source']['scan_range'] = _rectangle(c['source'].get('scan_range') or '', '表C扫描区域')
    c['upload']['scan_range'] = _rectangle(c['upload'].get('scan_range') or '', '表A扫描区域')

    # 下载表C成功后自动回填；本次启动可选择沿用已有文件。
    format_file = c['source'].get('format_file')
    if not isinstance(format_file, str):
        raise ValueError('表C格式参考路径必须为文本')
    format_file = format_file.strip()
    if format_file and Path(format_file).suffix.lower() != '.xlsx':
        raise ValueError('表C格式参考路径必须指向.xlsx文件')
    c['source']['format_file'] = format_file
    format_sheet = c['source'].get('format_sheet')
    if not isinstance(format_sheet, str):
        raise ValueError('表C格式参考工作表名称必须为文本')
    c['source']['format_sheet'] = format_sheet.strip()

    # 主键列：单列字母，默认D。
    pk = str(c['sync'].get('primary_key_column') or 'D').strip().upper()
    try:
        column(pk)
    except ValueError:
        raise ValueError('主键列须填写单个列字母，例如 D（默认第4列）')
    c['sync']['primary_key_column'] = pk
    c['sync']['skip_columns'] = format_skip_columns(sync_skip_columns(c))
    c['sync']['sync_columns'] = format_skip_columns(
        parse_sync_columns(c['sync'].get('sync_columns', '')))

    # 表A1路径。
    local_file = c['local'].get('local_file')
    if not isinstance(local_file, str):
        raise ValueError('本地表A1路径必须为文本')
    if local_file.strip() and Path(local_file).suffix.lower() != '.xlsx':
        raise ValueError('表A1路径必须指向.xlsx文件')
    c['local']['local_file'] = local_file.strip() or '表A1.xlsx'
    if len(c['local']['sheet_name']) > 31 or re.search(r'[\\/*?:\[\]]', c['local']['sheet_name']) or c['local']['sheet_name'] == '_sync_state':
        raise ValueError('表A1工作表名称无效或使用了保留名称')

    options = c['filter']
    if not isinstance(options['enabled'], bool):
        raise ValueError('筛选镜像开关必须为开关值')
    number('filter', 'header_row', 1, 1048576, True)
    for key in ('sheet_name', 'mirror_sheet'):
        name = options[key]
        if not isinstance(name, str) or not name.strip() or len(name.strip()) > 31 or re.search(r'[\\/*?:\[\]]', name) or name == '_sync_state':
            raise ValueError('筛选工作表名称无效')
        options[key] = name.strip()
    if not isinstance(options['local_file'], str) or not options['local_file'].strip():
        raise ValueError('请填写独立的筛选镜像路径')
    options['local_file'] = options['local_file'].strip()
    if Path(options['local_file']).suffix.lower() != '.xlsx':
        raise ValueError('筛选镜像路径必须指向.xlsx文件')
    if options['enabled']:
        from filtered_mirror import effective_config
        effective_config(ROOT, c)
        if options['sheet_name'] == c['source']['sheet_name']:
            raise ValueError('系统配置工作表必须与表C数据工作表不同')
        if options['header_row'] >= c['sync']['data_start_row']:
            raise ValueError('表C表头行必须位于数据起始行之前')

    json.dumps(c, allow_nan=False)
    return c


class App:
    def __init__(self, root):
        self.root = Path(root)
        self.token = secrets.token_urlsafe(32)
        self.lock = threading.RLock()
        self.reports = Reports(self.root)
        self.file_stamps = self.code_stamps()
        from edge_copy import EdgeCopier
        self.copier = EdgeCopier(self.root)
        self.copier.reports = self.reports
        from scheduler import TaskScheduler
        self.scheduler = TaskScheduler(self.copier, self.reports, self.config)
        try:
            self.scheduler.configure(self.config())
        except Exception:
            pass
        activity = getattr(self.copier, 'activity', None)
        if activity is not None:
            activity.record('网页服务启动', pid=os.getpid(), version=VERSION,
                            directory=str(self.root.resolve()))

    @staticmethod
    def code_stamps():
        paths = list(ROOT.glob('*.py')) + [ROOT / 'web' / 'index.html']
        return {str(p.relative_to(ROOT)): (p.stat().st_mtime_ns, p.stat().st_size)
                for p in paths if p.exists()}

    def version(self):
        return dict(version=VERSION, capabilities=CAPABILITIES, pid=os.getpid(),
                    program_directory=str(self.root.resolve()), executable=sys.executable,
                    restart_required=self.file_stamps != self.code_stamps())

    def release_notes(self):
        """网页底部的更新说明，来源：web/release_notes.json（新条目在前）。"""
        path = ROOT / 'web' / 'release_notes.json'
        try:
            data = json.loads(path.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            return []
        if isinstance(data, list):
            return data
        if isinstance(data, dict) and isinstance(data.get('entries'), list):
            return data['entries']
        return []

    def config(self):
        path = self.root / 'config.json'
        c = json.loads(path.read_text(encoding='utf-8')) if path.is_file() else {}
        migrated = _migrate_legacy(c)
        if migrated is not None:
            c = migrated
            self._write_config(c)
        if 'mysql' in c:
            c.pop('mysql')
            c['local'] = {'local_file': '表A1.xlsx', 'sheet_name': '表A1'}
            self._write_config(c)
        # 旧配置用 minutes（分钟），统一迁移成 wait_seconds（秒），并移除旧字段。
        sync = c.get('sync')
        if isinstance(sync, dict):
            if 'wait_seconds' not in sync:
                try:
                    sync['wait_seconds'] = max(1, int(round(float(sync.get('minutes', 2.0)) * 60)))
                except (TypeError, ValueError):
                    sync['wait_seconds'] = 120
            sync.pop('minutes', None)
        defaults = _defaults()
        for section, values in defaults.items():
            c.setdefault(section, {})
            for key, value in values.items():
                c[section].setdefault(key, value)
        c['sync'].pop('scan_rows', None)
        # 旧版“本地表B对比”开关已废弃：线下对比现在是默认且唯一的方式，移除残留字段。
        c['sync'].pop('use_table_b', None)
        return c

    def _write_config(self, c):
        path = self.root / 'config.json'
        temp = path.with_suffix('.web.tmp')
        temp.write_text(json.dumps(c, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temp, path)

    def save(self, c):
        with self.lock:
            validate_config(c)
            self._write_config(c)
            self.scheduler.configure(c)
            activity = getattr(self.copier, 'activity', None)
            if activity is not None:
                sync = c.get('sync', {})
                activity.record('保存配置', monitor=str(c.get('monitor', {}).get('enabled')),
                                primary_key=sync.get('primary_key_column'),
                                wait_seconds=sync.get('wait_seconds'),
                                batch_rows=sync.get('batch_rows'),
                                skip_columns=sync.get('skip_columns'),
                                sync_columns=sync.get('sync_columns'),
                                insert_enabled=sync.get('insert_enabled'),
                                watch_enabled=sync.get('watch_enabled'),
                                legacy_scan=sync.get('legacy_scan'),
                                max_pastes=sync.get('max_pastes'),
                                source_sheet=c.get('source', {}).get('sheet_name'),
                                upload_sheet=c.get('upload', {}).get('sheet_name'),
                                local_file=c.get('local', {}).get('local_file'),
                                filter_enabled=c.get('filter', {}).get('enabled'))

    def source_info(self):
        import datetime as dt
        from local_mirror import mirror_path
        path = mirror_path(self.root, self.config())
        try:
            stat = path.resolve().stat()
            return dict(ok=True, path=str(path), modified_at=dt.datetime.fromtimestamp(
                stat.st_mtime, dt.timezone.utc).isoformat())
        except (OSError, ValueError):
            return dict(ok=False)

    def sync(self):
        with self.lock:
            self.copier.submit_sync(validate_config(self.config()))

    def open_table(self, data):
        with self.lock:
            if not isinstance(data, dict) or set(data) != {'which'} or data['which'] not in ('c', 'a'):
                raise ValueError('打开操作参数无效')
            self.copier.submit_open(validate_config(self.config()), data['which'])

    def download_table_c(self):
        with self.lock:
            if self.scheduler.state()['running']:
                raise ValueError('请先停止定时任务，再手动下载表C')
            self.copier.submit_download_table_c(validate_config(self.config()))

    def download_table_b(self):
        with self.lock:
            if self.scheduler.state()['running']:
                raise ValueError('请先停止定时任务，再手动下载表B')
            self.copier.submit_download_table_b(validate_config(self.config()))

    def table_b_info(self):
        import datetime as dt
        path = self.root / 'DATA' / '表B.xlsx'
        try:
            stat = path.stat()
            return dict(ok=True, path=str(path.resolve()), size=stat.st_size,
                        modified_at=dt.datetime.fromtimestamp(stat.st_mtime, dt.timezone.utc).isoformat())
        except OSError:
            return dict(ok=False, path=str(path.resolve()))

    def table_c_info(self):
        import datetime as dt
        path = self.root / 'DATA' / '表C.xlsx'
        try:
            stat = path.stat()
            return dict(ok=True, path=str(path.resolve()), size=stat.st_size,
                        modified_at=dt.datetime.fromtimestamp(stat.st_mtime, dt.timezone.utc).isoformat())
        except OSError:
            return dict(ok=False, path=str(path.resolve()))

    def start(self, data=None):
        with self.lock:
            data = {} if data is None else data
            if (not isinstance(data, dict) or set(data) - {'download_table_c'} or
                    not isinstance(data.get('download_table_c', True), bool)):
                raise ValueError('本次启动是否下载表C必须为开关值')
            validate_config(self.config())
            self.scheduler.start(download_table_c=data.get('download_table_c', True))

    def stop(self):
        with self.lock:
            self.scheduler.stop()
            # 同时请求中断“正在执行”的当前同步轮次（run_sync 在下一检查点抛出 SyncStopped）。
            self.copier.request_stop()

    def state(self):
        with self.lock:
            edge = self.copier.state()
            return dict(running=self.scheduler.state()['running'],
                        schedule=self.scheduler.state(),
                        source_data=self.source_info(),
                        table_c=self.table_c_info(),
                        table_b=self.table_b_info(),
                        edge_copy=edge, upload=edge.get('upload', {}),
                        server=self.version(), latest_error=self.reports.latest())


def _basic_ok(header, expected):
    if not expected:
        return True
    if not header or not header.startswith('Basic '):
        return False
    import base64
    try:
        raw = base64.b64decode(header[6:].strip()).decode('utf-8')
    except Exception:
        return False
    return secrets.compare_digest(raw, expected)


def handler_for(app, *, allowed_hosts=None, allowed_origins=None, auth=None):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def fail(self, exc, status=400, code=None, message=None):
            from edge_copy import TaskBusyError
            task = exc.task if isinstance(exc, TaskBusyError) else None
            if task is not None:
                status = 409
            context = {'method': self.command, 'path': self.route, 'http_status': status}
            if task:
                context.update({'task_' + k: v for k, v in task.items()})
            report = app.reports.record(exc, 'HTTP ' + self.command + ' ' + self.route,
                                        code=code, message=message or str(exc), context=context)
            self.send(status, dict(error=report['message'], code=report['code'],
                                   report_id=report['id'], report=report, active_task=task,
                                   cell_limit=exc.details if isinstance(exc, CellLimitError) else None))

        @property
        def route(self):
            return urllib.parse.urlsplit(self.path).path.rstrip('/') or '/'

        def missing_route(self):
            self.fail(ValueError('接口不存在'), 404, 'API_NOT_FOUND',
                      '本地服务没有此接口：' + self.route + '。请确认补丁完整覆盖，并关闭旧服务窗口后从当前目录重启 start_ui.bat，再按 Ctrl+F5。')

        def send(self, code, body, kind='application/json; charset=utf-8'):
            try:
                self.send_body(code, body, kind)
            except ConnectionError:
                self.close_connection = True

        def send_body(self, code, body, kind):
            if not isinstance(body, bytes):
                body = json.dumps(body, ensure_ascii=False).encode('utf-8')
            self.send_response(code)
            self.send_header('Content-Type', kind)
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; frame-ancestors 'none'; base-uri 'none'")
            self.end_headers()
            self.wfile.write(body)

        def send_favicon(self):
            """Serve the static favicon from web/favicon.ico (binary, cacheable).

            旧实现返回 204 空响应，浏览器在无缓存的新电脑上拿不到图标；这里改为
            直接输出真实 .ico 文件，并允许缓存一天，避免每次导航都重新请求。
            """
            path = ROOT / 'web' / 'favicon.ico'
            try:
                data = path.read_bytes()
            except OSError:
                self.send(404, {'error': 'favicon 文件缺失', 'code': 'FAVICON_MISSING'})
                return
            try:
                self.send_response(200)
                self.send_header('Content-Type', 'image/x-icon')
                self.send_header('Content-Length', str(len(data)))
                self.send_header('Cache-Control', 'public, max-age=86400')
                self.send_header('X-Content-Type-Options', 'nosniff')
                self.end_headers()
                self.wfile.write(data)
            except ConnectionError:
                self.close_connection = True

        def unauthorized(self):
            body = json.dumps({'error': '需要登录（HTTP Basic 认证）', 'code': 'AUTH_REQUIRED'},
                              ensure_ascii=False).encode('utf-8')
            self.send_response(401)
            self.send_header('WWW-Authenticate', 'Basic realm="wps-sync"')
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(body)

        def authorized(self, api=False):
            if auth and not _basic_ok(self.headers.get('Authorization'), auth):
                self.unauthorized()
                return False
            port = self.server.server_port
            hosts = allowed_hosts or {f'127.0.0.1:{port}'}
            if self.headers.get('Host', '') not in hosts:
                self.fail(ValueError('仅允许本机地址访问'), 403, 'LOCAL_ONLY')
                return False
            origins = allowed_origins or {f'http://127.0.0.1:{port}'}
            origin = self.headers.get('Origin')
            if origin and origin not in origins:
                self.fail(ValueError('拒绝跨站请求'), 403, 'CROSS_ORIGIN_DENIED')
                return False
            if api and not secrets.compare_digest(self.headers.get('X-App-Token', ''), app.token):
                self.send(403, {'error': '页面连接已过期，正在重新连接；如果仍失败，请刷新网页后重试。', 'code': 'SESSION_EXPIRED'})
                return False
            return True

        def do_GET(self):
            if not self.authorized(self.route.startswith('/api/') and self.route not in ('/api/session', '/api/version')):
                return
            try:
                if self.route == '/':
                    html = (ROOT / 'web' / 'index.html').read_text(encoding='utf-8').replace('__APP_TOKEN__', app.token)
                    self.send(200, html.encode('utf-8'), 'text/html; charset=utf-8')
                elif self.route == '/favicon.ico':
                    self.send_favicon()
                elif self.route == '/api/session':
                    self.send(200, {'token': app.token})
                elif self.route == '/api/version':
                    self.send(200, app.version())
                elif self.route == '/api/diagnostics':
                    self.send(200, {'reports': app.reports.recent()})
                elif self.route == '/api/config':
                    self.send(200, app.config())
                elif self.route == '/api/status':
                    self.send(200, app.state())
                elif self.route == '/api/release-notes':
                    self.send(200, app.release_notes())
                else:
                    self.missing_route()
            except Exception as exc:
                self.fail(exc, 400 if isinstance(exc, ValueError) else 500)

        def do_POST(self):
            if not self.authorized(True):
                return
            try:
                if self.version_changed():
                    return
                length = int(self.headers.get('Content-Length', '0'))
                if not 0 < length <= 65536:
                    raise ValueError('请求过大')
                data = json.loads(self.rfile.read(length))
                if self.route == '/api/config':
                    app.save(data)
                elif self.route == '/api/sync':
                    app.sync()
                elif self.route == '/api/open':
                    app.open_table(data)
                elif self.route == '/api/start':
                    app.start(data)
                elif self.route == '/api/download-table-c':
                    if data != {}:
                        raise ValueError('下载表C无需额外参数')
                    app.download_table_c()
                elif self.route == '/api/download-table-b':
                    if data != {}:
                        raise ValueError('下载表B无需额外参数')
                    app.download_table_b()
                elif self.route == '/api/stop':
                    app.stop()
                else:
                    self.missing_route()
                    return
                self.send(200, {'ok': True})
            except Exception as exc:
                self.fail(exc, 400 if isinstance(exc, ValueError) else 500)

        def version_changed(self):
            if self.route != '/api/stop' and app.version()['restart_required']:
                self.fail(ValueError('程序文件已更新，当前后台仍是旧进程。请停止任务并关闭服务窗口，重新运行 start_ui.bat。'), 409, 'SERVER_RESTART_REQUIRED')
                return True
            return False

    return Handler


def _access_sets(port, extra_hosts):
    """默认只允许本机访问；额外域名/IP 需显式用 --allow-host 放开。"""
    hosts = {f'127.0.0.1:{port}', f'localhost:{port}'}
    origins = {f'http://127.0.0.1:{port}', f'http://localhost:{port}'}
    for raw in extra_hosts or []:
        value = str(raw).strip().lower().rstrip('/')
        if not value:
            continue
        if '://' in value:
            parts = urllib.parse.urlsplit(value)
            if parts.netloc:
                hosts.add(parts.netloc)
            origins.add(value)
            continue
        hosts.add(value)
        hosts.add(f'{value}:{port}')
        origins.add(f'http://{value}')
        origins.add(f'https://{value}')
        origins.add(f'http://{value}:{port}')
    return hosts, origins


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=8765)
    parser.add_argument('--host', default=os.environ.get('SYNC_HOST', '127.0.0.1'),
                        help='监听地址；远程访问用 0.0.0.0（默认仅本机 127.0.0.1）')
    parser.add_argument('--no-browser', action='store_true')
    parser.add_argument('--auto-start', action='store_true',
                        help='网页服务启动后立即启动定时任务（无人值守模式）')
    parser.add_argument('--allow-host', action='append', default=None,
                        help='额外允许的访问域名或 IP（可多次指定），需配合 --host 0.0.0.0')
    parser.add_argument('--auth', default=os.environ.get('SYNC_AUTH', ''),
                        help='user:password；设置后启用 HTTP Basic 认证')
    args = parser.parse_args()
    auth = str(args.auth or '').strip() or None
    if auth is not None and ':' not in auth:
        parser.error('--auth 须为 user:password 形式')
    allowed_hosts, allowed_origins = _access_sets(args.port, args.allow_host)
    app = App(ROOT)
    server = ThreadingHTTPServer((args.host, args.port), handler_for(
        app, allowed_hosts=allowed_hosts, allowed_origins=allowed_origins, auth=auth))
    print(f'版本：{VERSION}\n程序目录：{ROOT}\nPython：{sys.executable}\n'
          f'监听：{args.host}:{server.server_port}\n'
          f'本机界面地址：http://127.0.0.1:{server.server_port}；关闭终端会停止服务。', flush=True)
    if args.auto_start:
        try:
            app.start()
            print('已自动启动定时任务。', flush=True)
        except Exception as exc:
            print(f'自动启动定时任务失败：{exc}', flush=True)
    if not args.no_browser:
        webbrowser.open(f'http://127.0.0.1:{server.server_port}')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        app.stop()
        app.copier.close()
        server.server_close()


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        from diagnostics import error_code, Reports
        message = ('端口已被占用，可能仍有旧服务运行。请关闭原 start_ui.bat 的服务窗口后重试；不要只关闭网页。'
                   if error_code(exc) == 'PORT_IN_USE' else str(exc))
        report = Reports(ROOT).record(exc, '启动网页服务', message=message)
        print(f'启动失败 [{report["code"]}] {report["message"]}\n报告编号：{report["id"]}\n报告目录：{ROOT / "error_reports"}', flush=True)
        raise SystemExit(1)
