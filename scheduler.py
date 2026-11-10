"""整轮同步实际完成（或失败）后等待设置的秒数，再从表C起始位置扫描。"""
import threading
import time

TICK_SECONDS = 0.2
MIN_WAIT_SECONDS = 1.0


class TaskScheduler:
    def __init__(self, copier, reports, get_config):
        self.copier = copier
        self.reports = reports
        self.get_config = get_config
        self.lock = threading.Lock()
        self.running = False
        self.thread = None
        self.stop_event = threading.Event()
        self.seen_completion = None
        self.tasks = {'sync': dict(enabled=False, wait_seconds=MIN_WAIT_SECONDS,
                                   last=None, last_ok=None, last_error='',
                                   next_run_at=None, awaiting_completion=False)}

    def configure(self, config):
        with self.lock:
            sync = config.get('sync', {})
            item = self.tasks['sync']
            item['enabled'] = bool(sync.get('enabled', False))
            try:
                wait_seconds = int(sync.get('wait_seconds', 120))
            except (TypeError, ValueError):
                wait_seconds = 120
            # 其他人的修改不会产生“本地编辑事件”，只能靠每一轮全表扫描发现，
            # 因此这个等待时间决定别人改完多久后才会同步到表A。
            item['wait_seconds'] = max(MIN_WAIT_SECONDS, wait_seconds)
        # 把最新配置交给 copier，供 worker 空闲时“改动即同步”轮询使用。
        try:
            self.copier.set_sync_config(config)
        except Exception:
            pass

    def _log(self, action, **fields):
        activity = getattr(self.copier, 'activity', None)
        if activity is not None:
            activity.record(action, **fields)

    def start(self, download_table_c=True):
        with self.lock:
            if self.running:
                return
            # Set preparation before the worker can queue the first round.
            self.copier.mark_prepare(download_table_c=download_table_c)
            self.running = True
            self.stop_event = threading.Event()
            self.thread = threading.Thread(target=self._loop, args=(self.stop_event,), daemon=True)
            self.thread.start()
        self._log('启动定时任务', wait_seconds=self.tasks['sync'].get('wait_seconds'),
                  download_table_c=download_table_c)

    def stop(self):
        with self.lock:
            self.running = False
            self.stop_event.set()
        self._log('停止定时任务')

    def state(self):
        with self.lock:
            return dict(running=self.running,
                        tasks={key: dict(value) for key, value in self.tasks.items()})

    def _maybe_run(self, config, stop_event=None):
        with self.lock:
            if stop_event is not None and stop_event.is_set():
                return
            item = self.tasks['sync']
            state = self.copier.state()
            completed = state.get('last_sync')
            if completed and completed['id'] != self.seen_completion:
                self.seen_completion = completed['id']
                item['awaiting_completion'] = False
                item['next_run_at'] = completed['finished'] + item['wait_seconds']
                item['last_error'] = completed['error']
                if completed['success']:
                    item['last_ok'] = completed['completed_at']
            if not item['enabled'] or state['busy'] or item['awaiting_completion']:
                return
            if item['next_run_at'] is not None and time.monotonic() < item['next_run_at']:
                return
            try:
                self.copier.submit_sync(config)
            except Exception as exc:
                if type(exc).__name__ == 'TaskBusyError':
                    return
                item['next_run_at'] = time.monotonic() + item['wait_seconds']
                item['last_error'] = str(exc)[:200]
                self.reports.record(exc, '定时同步：表C→表A', message=str(exc))
                self._log('定时同步失败', error=str(exc)[:200])
                return
            item['last'] = time.monotonic()
            # 排队成功不代表执行完成；必须收到工作线程的完成记录。
            item['awaiting_completion'] = True

    def _loop(self, stop_event):
        while not stop_event.is_set():
            try:
                config = self.get_config()
                self.configure(config)
                self._maybe_run(config, stop_event)
            except Exception as exc:
                with self.lock:
                    self.tasks['sync']['last_error'] = str(exc)[:200]
            stop_event.wait(TICK_SECONDS)
