"""专用 Edge 浏览器管理：读取在线表C、写入在线表A，走单一任务队列与忙锁。

所有 Playwright 对象都运行在同一个工作线程上；不等待网页请求返回 UI。
扫描表C与写入表A共用同一个专用 Edge（登录共享），但使用独立的标签页。
"""
import copy
import datetime as dt
import json
from pathlib import Path
import queue
import threading
import time
import urllib.parse
import uuid

from diagnostics import Reports, error_code
from watch import INSTALL_JS as WATCH_JS, install as _install_watch, drain_dirty_rows as _drain_dirty_rows


class TaskBusyError(ValueError):
    code = 'TASK_BUSY'

    def __init__(self, task):
        self.task = task or {}
        name = self.task.get('name', '同步任务')
        phase = self.task.get('phase', '正在处理')
        super().__init__(f'当前正在执行“{name}”：{phase}。请等待当前任务结束，本次没有重复提交。')


def validate_url(value, label='分享链接'):
    url = urllib.parse.urlparse(value)
    if url.scheme != 'https' or url.hostname not in {'www.kdocs.cn', 'kdocs.cn'} \
            or not url.path.startswith('/l/') or url.username or url.password \
            or url.port not in (None, 443):
        raise ValueError(f'请先填写有效的{label}并保存配置')
    return value


def _merge_insert_result(result, insert):
    """把“新增记录插入”的统计并入整行同步结果，并拼接提示信息。"""
    merged = dict(result)
    merged['inserted'] = insert.get('inserted', 0)
    merged['insert_writes'] = insert.get('writes', 0)
    merged['insert_overwritten'] = insert.get('overwritten', 0)
    merged['insert_header_skipped'] = insert.get('header_skipped', 0)
    merged['insert_unmappable'] = insert.get('unmappable', 0)
    if insert.get('inserted') or insert.get('header_skipped') or insert.get('unmappable'):
        merged['message'] = str(merged.get('message', '')) + ' ' + str(insert.get('message', ''))
    return merged


class EdgeCopier:
    def __init__(self, root):
        self.root = root
        self.reports = Reports(self.root)
        from activity_log import ActivityLog
        self.activity = ActivityLog(self.root)
        self.lock = threading.Lock()
        self.jobs = queue.Queue()
        self.status = dict(busy=False, ready=False, message='首次使用请先打开专用Edge。')
        self.upload_status = dict(busy=False, message='尚未同步表A。')
        self.thread = None
        self.context = self.page = self.playwright = None
        self.share_url = None
        self.page_c = None
        self.page_c_url = None
        self.page_a = None
        self.page_a_url = None
        self.active_task = None
        self.last_sync = None
        self._task_start = self._phase_start = 0
        self._context_closed = False
        self._watch_probe_at = 0.0
        self._watch_probe_cache = None
        self._watch_install_at = 0.0
        self._watch_config = None
        self._watch_pending = set()
        self._watch_changed_at = None
        self._stop_event = threading.Event()
        self._overwatch_built = False
        self._prepare_needed = False
        self._prepare_download_table_c = True
        self._table_b_fresh = False

    def _task_state(self):
        if self.active_task is None:
            return None
        return dict(self.active_task, elapsed_seconds=int(time.monotonic() - self._task_start),
                    unchanged_seconds=int(time.monotonic() - self._phase_start),
                    worker_alive=bool(self.thread and self.thread.is_alive()))

    def _refresh_watch_probe(self):
        """worker 线程：刷新“改动即同步”诊断缓存（约3秒一次，避免跨线程访问页面）。"""
        now = time.monotonic()
        if now - self._watch_probe_at < 3.0:
            return
        page = self.page_c
        result = None
        if page is not None:
            try:
                if not page.is_closed():
                    from watch import probe as _probe
                    result = _probe(page)
            except Exception as exc:
                result = [{'error': type(exc).__name__}]
        self._watch_probe_at = now
        self._watch_probe_cache = result

    def watch_probe(self):
        """返回最近一次 worker 刷新出的诊断结果（无 Playwright 调用，线程安全）。"""
        return self._watch_probe_cache

    def state(self):
        probe = self.watch_probe()
        with self.lock:
            return dict(copy.deepcopy(self.status), upload=copy.deepcopy(self.upload_status),
                        active_task=self._task_state(), last_sync=copy.deepcopy(self.last_sync),
                        watch_probe=probe)

    def _progress(self, values):
        if self.active_task is not None and 'message' in values and values['message'] != self.active_task['phase']:
            self.active_task['phase'] = values['message']
            self.active_task['updated_at'] = dt.datetime.now(dt.timezone.utc).isoformat()
            self._phase_start = time.monotonic()

    def update(self, **values):
        with self.lock:
            self.status.update(values)
            self._progress(values)

    def update_upload(self, **values):
        with self.lock:
            self.upload_status.update(values)
            self._progress(values)

    def _enqueue(self, job, name, message, *, upload=False):
        payload = copy.deepcopy(job)
        with self.lock:
            if self.status['busy']:
                raise TaskBusyError(self._task_state())
            self.status = dict(busy=True, ready=self.status['ready'], message=message)
            if upload:
                self.upload_status = dict(busy=True, message=message)
            now = dt.datetime.now(dt.timezone.utc).isoformat()
            self.active_task = dict(id=uuid.uuid4().hex[:12], name=name, kind=job[0],
                                    phase=message, started_at=now, updated_at=now)
            self._task_start = self._phase_start = time.monotonic()
            self.jobs.put(payload)
            try:
                if self.thread is None or not self.thread.is_alive():
                    self.thread = threading.Thread(target=self._run, daemon=True)
                    self.thread.start()
            except Exception:
                self.jobs.get_nowait()
                self.thread = None
                self.status['busy'] = False
                self.upload_status['busy'] = False
                self.active_task = None
                raise
        activity = getattr(self, 'activity', None)
        if activity is not None:
            activity.record('任务排队', kind=job[0], name=name, phase=message)

    def _finish_task(self, success=False):
        with self.lock:
            if self.active_task and self.active_task['kind'] in ('sync', 'sync_rows'):
                self.last_sync = dict(id=self.active_task['id'], success=success,
                                      finished=time.monotonic(),
                                      completed_at=dt.datetime.now(dt.timezone.utc).isoformat(),
                                      error='' if success else self.upload_status.get('message', '同步失败'))
            self.status['busy'] = self.upload_status['busy'] = False
            self.active_task = None

    def set_sync_config(self, config):
        """保存最新配置，供 worker 空闲时的“改动即同步”轮询使用。"""
        self._watch_config = copy.deepcopy(config)
        activity = getattr(self, 'activity', None)
        if activity is not None:
            activity.configure(config)

    def request_stop(self):
        """请求停止当前同步轮次：run_sync 在下一检查点抛出 SyncStopped。"""
        self._stop_event.set()
        activity = getattr(self, 'activity', None)
        if activity is not None:
            activity.record('请求停止', '停止后续轮次，当前轮次在下一检查点中断')

    def should_stop(self):
        return self._stop_event.is_set()

    def _reset_stop(self):
        self._stop_event.clear()

    def mark_prepare(self, download_table_c=True):
        """标记下一次同步前需要先执行“准备”：下载表C为表C并建立表A1镜像。

        由 TaskScheduler.start() 在每次启动定时任务后调用；准备成功后自动清除标记。"""
        self._prepare_download_table_c = download_table_c
        self._prepare_needed = True

    def submit_download_table_c(self, config):
        validate_url(config['source']['share_url'], '表C分享链接')
        self._reset_stop()
        self._enqueue(('download_c', config), '下载表C',
                      '正在准备下载表C到 DATA/表C.xlsx…', upload=True)

    def submit_download_table_b(self, config):
        validate_url(config['upload']['share_url'], '表A分享链接')
        self._reset_stop()
        self._enqueue(('download_b', config), '下载表B',
                      '正在准备下载表A到 DATA/表B.xlsx…', upload=True)

    def submit_sync(self, config):
        """按批扫描表C，与表A按主键比较并合并复制连续差异行。"""
        validate_url(config['source']['share_url'], '表C分享链接')
        validate_url(config['upload']['share_url'], '表A分享链接')
        self.set_sync_config(config)
        self._reset_stop()
        self._enqueue(('sync', config), '同步表C→表A',
                      '正在扫描表C…', upload=True)

    def submit_sync_rows(self, config, rows):
        """只同步指定的若干行（“改动即同步”的即时路径）。"""
        validate_url(config['source']['share_url'], '表C分享链接')
        validate_url(config['upload']['share_url'], '表A分享链接')
        rows = sorted({int(r) for r in rows
                       if isinstance(r, (int, float)) and 1 <= int(r) <= 1048576})
        if not rows:
            return
        self._reset_stop()
        self._enqueue(('sync_rows', config, rows), '同步表C→表A（改动即同步）',
                      f'正在同步刚修改的{len(rows)}行…', upload=True)

    def drain_dirty_rows(self):
        """从表C页面读取并清空被改动的行号（页面未打开时返回空列表）。"""
        page = self.page_c
        if page is None:
            return []
        try:
            if page.is_closed():
                return []
        except Exception:
            return []
        now = time.monotonic()
        if now - self._watch_install_at >= 3.0:
            self._watch_install_at = now
            _install_watch(page)  # 幂等；补装之后才挂载的 iframe
        try:
            return _drain_dirty_rows(page)
        except Exception:
            return []

    def submit_open(self, config, which):
        """打开表C或表A的专用 Edge 标签页，便于登录。"""
        if which not in ('c', 'a'):
            raise ValueError('打开操作无效')
        if which == 'c':
            validate_url(config['source']['share_url'], '表C分享链接')
            self._enqueue(('open_c', config), '打开表C', '正在打开表C…')
        else:
            validate_url(config['upload']['share_url'], '表A分享链接')
            self._enqueue(('open_a', config), '打开表A', '正在打开表A…', upload=True)

    # ---- 页面管理 ----
    def _ensure_context(self):
        if self.context is not None:
            try:
                if self._context_closed or (self.context.browser and not self.context.browser.is_connected()):
                    self._discard_context()
            except Exception:
                self._discard_context()
        if self.context is None:
            try:
                from playwright.sync_api import sync_playwright
            except ModuleNotFoundError as exc:
                raise ValueError('缺少Edge组件，请运行安装依赖.bat安装playwright后重试') from exc
            self.playwright = sync_playwright().start()
            try:
                self.context = self.playwright.chromium.launch_persistent_context(
                    str(self.root / 'edge_copy_profile'), channel='msedge', headless=False,
                    accept_downloads=True, no_viewport=True)
                # 每个页面（含刷新后）都注入“改动即同步”的监听脚本。
                self.context.add_init_script(WATCH_JS)
                self._context_closed = False
                context = self.context

                def closed(*_args):
                    if self.context is context:
                        self._context_closed = True
                context.on('close', closed)
            except Exception as exc:
                self.playwright.stop()
                self.playwright = None
                raise ValueError('无法启动Microsoft Edge，请确认已安装Edge，并关闭之前的同步专用Edge后重试') from exc

    def _discard_context(self):
        self.context = self.page = self.page_c = self.page_a = None
        self.share_url = self.page_c_url = self.page_a_url = None
        self._context_closed = False
        driver, self.playwright = self.playwright, None
        if driver is not None:
            try:
                driver.stop()
            except Exception:
                pass

    def _open_page(self, url, page_attr, url_attr, *, reload=False, label=''):
        """打开（或复用）一个在线表格标签页；返回 page。"""
        for attempt in range(2):
            try:
                self._ensure_context()
                page = getattr(self, page_attr)
                if page is None or page.is_closed():
                    page = self.context.new_page()
                    setattr(self, page_attr, page)
                    setattr(self, url_attr, None)
                current = None
                try:
                    current = urllib.parse.urlsplit(page.url)
                except Exception:
                    current = None
                target = urllib.parse.urlsplit(url)
                already_there = bool(current) and current.netloc == target.netloc and current.path == target.path
                if getattr(self, url_attr) != url and not (reload and already_there):
                    page.goto(url, wait_until='domcontentloaded', timeout=60000)
                    setattr(self, url_attr, url)
                elif reload:
                    setattr(self, url_attr, url)
                return page
            except Exception as exc:
                closed = type(exc).__name__ == 'TargetClosedError'
                if not closed or attempt:
                    if closed:
                        raise ValueError(f'打开{label}失败：专用Edge连接已关闭。请重新点击“打开{label}”后重试') from exc
                    raise ValueError(f'打开{label}失败：' + str(exc)[:160] + '。请稍后重试') from exc
                self._discard_context()
        raise ValueError(f'打开{label}失败')

    def _page_c(self, url, *, reload=False):
        return self._open_page(url, 'page_c', 'page_c_url', reload=reload, label='表C')

    def _page_a(self, url, *, reload=False):
        return self._open_page(url, 'page_a', 'page_a_url', reload=reload, label='表A')

    # ---- 剪贴板读取与选区 ----
    @staticmethod
    def _blank_clip(ref, sequence=None):
        from copy_range import column_name
        from update_excel import address
        from excel_clipboard import read_clipboard
        left, right = ref.split(':')
        c1, r1 = address(left)
        c2, r2 = address(right)
        rows, cols = r2 - r1 + 1, c2 - c1 + 1
        text = ''.join('\t' * (cols - 1) + '\r\n' for _ in range(rows))
        html = '<table>' + ''.join('<tr>' + '<td></td>' * cols + '</tr>' for _ in range(rows)) + '</table>'
        if sequence is None:
            sequence = read_clipboard()['sequence']
        return dict(text=text, html=html, sequence=sequence, blank=True)

    def _set_local_clipboard(self, rows, clip):
        from local_clipboard import write_clipboard, verify_clipboard
        write_clipboard(clip)
        clip['sequence'] = verify_clipboard(clip)

    def _verify_local_clipboard(self, clip):
        from local_clipboard import verify_clipboard
        verify_clipboard(clip)

    def _cell_format_source(self, config):
        from html_cell_formats import HtmlCellFormats
        return HtmlCellFormats(self, config)

    def _page_label(self, page):
        if page is self.page_c:
            return '表C'
        if page is self.page_a:
            return '表A'
        return '页面'

    def _op_log(self, action, page, ref='', **fields):
        """把单元格级操作（选择/复制/粘贴）写入操作日志；日志开关关闭时不产生记录。"""
        activity = getattr(self, 'activity', None)
        if activity is None:
            return
        try:
            label = self._page_label(page)
        except Exception:
            label = '页面'
        activity.record(action, page=label, ref=(ref or ''), **fields)

    def _paste(self, page, ref, values=None):
        start = time.monotonic()
        try:
            page.keyboard.press('Control+v')
        finally:
            self._op_log('粘贴单元格', page, ref, values=self._paste_rows(values),
                         耗时=self._elapsed_text(start))

    @staticmethod
    def _paste_rows(values):
        """把粘贴内容规整成「行列表」；单行列表包装成一行，空内容返回 None（日志省略该字段）。"""
        if not values:
            return None
        if isinstance(values, (list, tuple)) and values and isinstance(values[0], (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]

    @staticmethod
    def _elapsed_text(start):
        """把起始 monotonic 时间换算成「秒」文本，用于操作日志的耗时字段。"""
        return f'{time.monotonic() - start:.3f}s'

    def _capture(self, page, ref):
        page.bring_to_front()
        marker = '__wps_copy_' + uuid.uuid4().hex
        page.evaluate('(text)=>navigator.clipboard.writeText(text)', marker)
        self._select(page, ref or 'A1')
        if ref is None:
            page.keyboard.press('Control+Home')
            page.keyboard.press('Control+Shift+End')
        start = time.monotonic()
        try:
            page.keyboard.press('Control+c')
            from excel_clipboard import read_clipboard, CopyIntegrityError
            clip = None
            plain = marker
            blank_samples = 0
            blank_sequence = None
            for attempt in range(160):
                if attempt and attempt % 15 == 0:
                    page.bring_to_front()
                    try:
                        page.keyboard.press('Control+c')
                    except Exception:
                        pass
                try:
                    candidate = read_clipboard()
                except CopyIntegrityError:
                    raise
                except Exception:
                    page.wait_for_timeout(200)
                    continue
                plain = candidate.get('text') or ''
                if plain != marker and candidate['html']:
                    clip = candidate
                    break
                if plain != marker and not candidate.get('html') and not plain.strip():
                    # WPS may publish empty text before rich HTML. Give that HTML
                    # time to arrive; a value-only placeholder cannot prove styles.
                    if blank_sequence != candidate['sequence']:
                        blank_samples = 0
                    blank_sequence = candidate['sequence']
                    blank_samples += 1
                    if blank_samples >= 4:
                        break
                else:
                    blank_samples = 0
                page.wait_for_timeout(200)
            if clip is None:
                if ref and not plain.strip():
                    self.update(message=f'{ref} 复制只返回空白值，未返回格式；此结果仅用于内容比对…')
                    return self._blank_clip(ref, sequence=blank_sequence)
                if plain not in ('', marker) and not (clip or {}).get('html'):
                    raise ValueError('网页只给出了纯文本、没有带格式的表格，请缩小扫描区域或每批行数')
                raise ValueError('未取得剪贴板表格：复制没有生效，请检查登录、表格定位和复制权限')
            return clip
        finally:
            self._op_log('复制单元格', page, ref or '整表', 耗时=self._elapsed_text(start))

    def _used_range(self, page):
        """Discover bounds through Ctrl+End without copying excluded columns."""
        from range_selection import find_namebox, _dispose
        from update_excel import address
        page.bring_to_front()
        self._select(page, 'A1')
        page.keyboard.press('Control+End')
        page.wait_for_timeout(300)
        _frame, item, detail = find_namebox(page)
        try:
            last = str(detail.get('address') or '').replace('$', '').upper().split(':')[-1]
            address(last)
            return f'A1:{last}'
        finally:
            _dispose(item)

    def _select(self, page, ref):
        start = time.monotonic()
        try:
            from range_selection import select_range
            from paste_menu import dismiss_paste_options
            dismiss_paste_options(page)

            def progress(text):
                task = self.state().get('active_task') or {}
                if task.get('kind') in ('sync', 'open_a', 'upload'):
                    self.update_upload(message=text)
                else:
                    self.update(message=text)
            return select_range(page, ref, progress=progress)
        finally:
            self._op_log('选择单元格', page, ref, 耗时=self._elapsed_text(start))

    # ---- 任务执行 ----
    def _write_config_file(self, config):
        """把最新配置原子写回 config.json（用于回填表C路径）。"""
        import os
        path = Path(self.root) / 'config.json'
        temp = path.with_suffix('.web.tmp')
        temp.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temp, path)

    def _run_download_table_c(self, config):
        """Manual and startup downloads share the same worker-only operation."""
        from sync_runner import download_table_c
        activity = getattr(self, 'activity', None)
        if activity is not None:
            activity.record('开始下载表C', path='DATA/表C.xlsx')
        start = time.monotonic()
        table_c = download_table_c(self, config)
        if activity is not None:
            activity.record('表C下载完成', path=table_c.get('path'), sheet=table_c.get('sheet'),
                            size=table_c.get('size'), 耗时=self._elapsed_text(start))
        config = copy.deepcopy(config)
        config['source']['format_file'] = table_c['path']
        config['source']['format_sheet'] = table_c['sheet']
        try:
            self._write_config_file(config)
        except Exception as exc:
            raise ValueError('表C已下载到 DATA/表C.xlsx，但格式参考配置保存失败，请检查config.json的写入权限') from exc
        self.set_sync_config(config)
        # 方法2：build_mirror 建镜像时可直接复用刚下载的表C文件，不必再下载一次。
        self._table_c_fresh = True
        self.update_upload(message=table_c.get('message', '表C下载完成。'))
        return config

    def _run_download_table_b(self, config):
        """下载表A为本地表B（DATA/表B.xlsx）；方法照搬表C下载。"""
        from table_c_download import download_table_b
        activity = getattr(self, 'activity', None)
        if activity is not None:
            activity.record('开始下载表B', path='DATA/表B.xlsx')
        start = time.monotonic()
        table_b = download_table_b(self, config)
        if activity is not None:
            activity.record('表B下载完成', path=table_b.get('path'), sheet=table_b.get('sheet'),
                            size=table_b.get('size'), 耗时=self._elapsed_text(start))
        # 表B 是表A 的本地快照，用于与表C 本地对比差异；本轮标记为已下载，
        # 供 run_sync 复用（避免同一轮重复下载表A）。
        self._table_b_fresh = True
        self.update_upload(message=table_b.get('message', '表B下载完成。'))
        return table_b

    def _run_prepare(self, config):
        """Apply this start's download choice, then prepare the A1 mirror."""
        from sync_runner import build_mirror
        activity = getattr(self, 'activity', None)
        if activity is not None:
            activity.record('开始准备', download_table_c=self._prepare_download_table_c)
        if self._prepare_download_table_c:
            config = self._run_download_table_c(config)
            # If mirror preparation fails, retries need not download again.
            self._prepare_download_table_c = False
        elif activity is not None:
            activity.record('跳过表C下载', '本次启动沿用已有表C')
        if config.get('filter', {}).get('enabled'):
            return config  # Each filtered round prepares its own snapshot.
        self.update_upload(message='正在检查并建立表A1镜像…')
        build_mirror(self, config)
        if activity is not None:
            activity.record('镜像已建立', path=str(config.get('local', {}).get('local_file')),
                            sheet=str(config.get('local', {}).get('sheet_name')))
        return config

    def _run_sync(self, config):
        from overwatch import OverWatcher
        # 方法2：每轮“sync”任务清空表C文件快照，确保整表数据是本轮新下载的；
        # 同一轮内 build_mirror 与 run_sync 共享同一次下载（read_source_file 缓存）。
        self._source_snapshot = None
        self._table_b_snapshot = None
        self._table_b_merges = None
        # 每次启动定时任务后，按本次选择下载表C，并准备表A1镜像；
        # 之后每轮只做“整体扫描→比对→按主键写回表A”，不再重建镜像。
        if getattr(self, '_prepare_needed', False):
            config = self._run_prepare(config)
            self._prepare_needed = False
        if config.get('filter', {}).get('enabled'):
            from filtered_mirror import run_filtered
            return run_filtered(self, config)
        result = OverWatcher(self, config).run_once()
        if config.get('sync', {}).get('insert_enabled') and result.get('missing', 0):
            # 新增记录必须写回：不再因为“本轮刚粘贴过”而跳过插入。
            # run_insert 只写入表A 中不存在该主键的空白目标行（按锚点定位），
            # 与守护扫描的整行覆盖互不冲突；跳过会导致新记录永远是空白行。
            from cell_runner import run_insert
            result = _merge_insert_result(result, run_insert(self, config))
        return result

    def _run_sync_rows(self, config, rows):
        if config.get('filter', {}).get('enabled'):
            # A rule edit or membership change must never bypass a fresh filter.
            return self._run_sync(config)
        from overwatch import OverWatcher
        if not OverWatcher(self, config).ready():
            return OverWatcher(self, config).build()
        from sync_runner import run_sync
        result = run_sync(self, config, only_rows=rows)
        if config.get('sync', {}).get('insert_enabled') and result.get('missing', 0):
            # 新增记录必须写回：不再因为“本轮刚粘贴过”而跳过插入，否则新记录永远是空白行。
            from cell_runner import run_insert
            result = _merge_insert_result(result, run_insert(self, config, only_rows=rows))
        return result

    def _record_result(self, action, result, 耗时=None):
        """把一轮同步/即时同步的关键统计写入操作日志。"""
        activity = getattr(self, 'activity', None)
        if activity is None or not isinstance(result, dict):
            return
        keep = ('scanned', 'uploaded', 'modified', 'changed', 'writes', 'missing', 'blank',
                'paste_batches', 'pending', 'inserted', 'insert_writes', 'insert_overwritten',
                'filter_matched', 'filter_scanned')
        summary = {key: result[key] for key in keep if result.get(key) is not None}
        if 耗时 is not None:
            summary['耗时'] = 耗时
        activity.record(action, message=str(result.get('message', '')), **summary)

    def _poll_watch(self):
        """worker 线程空闲时轮询“改动即同步”的脏行并触发即时同步。

        所有 Playwright 调用都在 worker 线程内执行，避免跨线程访问页面。
        """
        self._refresh_watch_probe()  # 刷新诊断缓存（无论是否启用，便于 UI 提示）
        config = self._watch_config
        if not config or not config.get('sync', {}).get('watch_enabled'):
            return
        with self.lock:
            if self.status['busy']:
                return
        page = self.page_c
        if page is None:
            return
        try:
            if page.is_closed():
                return
        except Exception:
            return
        rows = self.drain_dirty_rows()
        now = time.monotonic()
        if rows:
            self._watch_pending.update(rows)
            self._watch_changed_at = now
        if not self._watch_pending or self._watch_changed_at is None:
            return
        debounce = max(0.0, float(config.get('sync', {}).get('watch_debounce') or 1.0))
        if now - self._watch_changed_at < debounce:
            return
        pending = sorted(self._watch_pending)
        self._watch_pending.clear()
        self._watch_changed_at = None
        activity = getattr(self, 'activity', None)
        if activity is not None:
            activity.record('检测到表C改动', rows=len(pending),
                            row_list=str(pending[:50]) + ('…' if len(pending) > 50 else ''))
        self.submit_sync_rows(config, pending)

    def _run(self):
        from sync_runner import SyncStopped
        try:
            while True:
                try:
                    job = self.jobs.get(timeout=0.5)
                except queue.Empty:
                    self._poll_watch()
                    continue
                if job is None:
                    break
                kind = job[0]
                success = False
                try:
                    if kind == 'sync':
                        config = job[1]
                        start = time.monotonic()
                        result = self._run_sync(config)
                        success = True
                        self._record_result('同步完成', result, 耗时=self._elapsed_text(start))
                        self.update_upload(result=result, message=result.get('message', '同步完成。'),
                                           completed_at=dt.datetime.now(dt.timezone.utc).isoformat())
                    elif kind == 'sync_rows':
                        config, rows = job[1], job[2]
                        start = time.monotonic()
                        result = self._run_sync_rows(config, rows)
                        success = True
                        self._record_result('即时同步完成', result, 耗时=self._elapsed_text(start))
                        self.update_upload(result=result, message=result.get('message', '同步完成。'),
                                           completed_at=dt.datetime.now(dt.timezone.utc).isoformat())
                    elif kind == 'download_c':
                        self._run_download_table_c(job[1])
                        success = True
                    elif kind == 'download_b':
                        self._run_download_table_b(job[1])
                        success = True
                    elif kind == 'open_c':
                        config = job[1]
                        page = self._page_c(config['source']['share_url'])
                        page.bring_to_front()
                        _install_watch(page)
                        self.update(ready=True, message='表C已打开。请自行登录并确认有编辑权限后，再执行同步。')
                        activity = getattr(self, 'activity', None)
                        if activity is not None:
                            activity.record('表C已打开', sheet=str(config.get('source', {}).get('sheet_name')))
                    elif kind == 'open_a':
                        config = job[1]
                        self._page_a(config['upload']['share_url']).bring_to_front()
                        self.update_upload(message='表A已打开。请自行登录并确认有编辑权限后，再执行同步。')
                        activity = getattr(self, 'activity', None)
                        if activity is not None:
                            activity.record('表A已打开', sheet=str(config.get('upload', {}).get('sheet_name')))
                except SyncStopped:
                    activity = getattr(self, 'activity', None)
                    if activity is not None:
                        activity.record('同步已停止', kind=kind)
                    self.update_upload(message='已停止：本轮同步已中止，已完成的部分（含表A1镜像）保留。')
                except Exception as exc:
                    code = error_code(exc)
                    if isinstance(exc, ValueError):
                        message = str(exc)
                    elif isinstance(exc, MemoryError):
                        message = '电脑可用内存不足，请缩小扫描区域或关闭其他程序后重试'
                    else:
                        message = ('表C下载失败，请检查表C登录状态、下载权限和网络。' if kind == 'download_c' else
                                   '表B下载失败，请检查表A登录状态、下载权限和网络。' if kind == 'download_b' else
                                   '打开表C失败，请检查表C登录状态和网络。' if kind == 'open_c' else
                                   '打开表A失败，请检查表A登录状态和网络。' if kind == 'open_a' else
                                   '同步失败，请检查登录、工作表和编辑权限；本轮变化未确认写入。')
                    report = self.reports.record(exc, '下载表C' if kind == 'download_c' else
                                                 '下载表B' if kind == 'download_b' else
                                                 '打开表C' if kind == 'open_c' else
                                                 '打开表A' if kind == 'open_a' else '同步表C→表A', message=message,
                                                 context={'progress': self.state().get('message', '')})
                    activity = getattr(self, 'activity', None)
                    if activity is not None:
                        activity.record('表C下载失败' if kind == 'download_c' else
                                        '表B下载失败' if kind == 'download_b' else
                                        '打开表C失败' if kind == 'open_c' else
                                        '打开表A失败' if kind == 'open_a' else '同步失败', kind=kind, code=report['code'],
                                        message=report['message'])
                    self.update_upload(error=report['message'], message=report['message'],
                                       code=report['code'], report_id=report['id'], report=report)
                finally:
                    self._finish_task(success)
        finally:
            if self.context is not None:
                try:
                    self.context.close()
                except Exception:
                    pass
            if self.playwright is not None:
                try:
                    self.playwright.stop()
                except Exception:
                    pass

    def close(self):
        if self.thread:
            self.jobs.put(None)
            self.thread.join(timeout=5)
