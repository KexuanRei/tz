"""Download table C's original XLSX through visible WPS export controls."""
import os
from pathlib import Path
import tempfile
import time

from online_format import OnlineFormatter

EXCEL_LABELS = ('导出为 Excel(.xlsx)', '导出为Excel(.xlsx)',
                '导出为 Excel (.xlsx)', '导出为Excel（.xlsx）',
                'Excel (.xlsx)', 'Excel（.xlsx）', 'Excel文件(.xlsx)',
                'Microsoft Excel (.xlsx)', 'Excel', '.xlsx', 'XLSX',
                '导出为Excel', '导出为Excel文件', '另存为Excel文件', '下载为Excel')
EXPORT_LABELS = ('导出为', '下载为', '导出', '下载到本地', '下载', '另存为')


def _main_menu(ui):
    for label in ('文件', 'File', '文件菜单', '主菜单', '菜单', 'Menu'):
        control = _control(ui, label)
        if control is not None:
            return control
    # The compact WPS skin has an unlabelled hamburger instead of "文件".
    # Require menu semantics, a compact top-left location, and a unique target;
    # never click a hard-coded screen coordinate or an arbitrary toolbar icon.
    found = []
    for frame in ui._frames():
        paths = frame.locator('[class*="menu" i],[class*="hamburger" i],'
                              '[aria-haspopup="menu"]').evaluate_all(r'''elements => {
            const targets = new Set();
            for (const el of elements) {
                const node = el.closest('button,[role="button"],[aria-haspopup="menu"]') || el;
                if (node.closest('[role="grid"],[role="gridcell"],td,th,[role="dialog"],[role="menu"]')) continue;
                const box = node.getBoundingClientRect(), style = getComputedStyle(node);
                if (style.visibility !== 'visible' || style.display === 'none' ||
                    box.x < 0 || box.x > 100 || box.y < 0 || box.y > 100 ||
                    box.width < 12 || box.width > 64 || box.height < 12 || box.height > 64) continue;
                if (!['', '☰', '≡'].includes((node.innerText || '').trim())) continue;
                targets.add(node);
            }
            // A menu wrapper and its nested icon describe one target.
            return [...targets].filter(node => ![...targets].some(other => other !== node && other.contains(node)))
                .map(node => {
                    const parts = [];
                    for (let el = node; el && el.nodeType === 1; el = el.parentElement) {
                        let i = 1;
                        for (let prev = el.previousElementSibling; prev; prev = prev.previousElementSibling)
                            if (prev.tagName === el.tagName) i++;
                        parts.unshift(el.tagName.toLowerCase() + ':nth-of-type(' + i + ')');
                    }
                    return parts.join(' > ');
                });
        }''')
        found.extend(frame.locator(path) for path in paths)
    if len(found) > 1:
        raise ValueError('表C左上角文件菜单不唯一，已停止下载')
    if found and not found[0].is_enabled():
        raise ValueError('表C左上角文件菜单不可用，请检查权限')
    return found[0] if found else None


def _expand_export(ui, control, label):
    # WPS's "导出为" is a hover submenu. Clicking again can close it.
    control.hover(timeout=3000)
    ui.page.wait_for_timeout(250)
    if any(_control(ui, excel) is not None for excel in EXCEL_LABELS):
        return
    # Retain support for skins whose submenu opens only on click.
    ui._click(control, label)


def _control(ui, label, scope=None):
    item = ui._unique_control((label,), scope=scope, optional=True)
    if item is not None:
        return item
    # Some WPS menus render div/span labels without ARIA roles. Only accept
    # exact text inside a visible menu/control, never a spreadsheet cell.
    found = []
    for root in ([scope] if scope is not None else ui._frames()):
        for candidate in root.get_by_text(label, exact=True).all():
            if candidate.is_visible() and candidate.evaluate('''el =>
                !el.closest('[role="grid"],[role="gridcell"],td,th') &&
                !!el.closest('button,[role="button"],[role="menu"],[role="menuitem"],'
                    + '[class*="menu" i],[role="dialog"]')'''):
                found.append(candidate)
    if len(found) > 1:
        raise ValueError('表C下载菜单中“' + label + '”不唯一，已停止下载')
    if found:
        if not found[0].is_enabled():
            raise ValueError('表C的“' + label + '”不可用，请检查下载权限')
        return found[0]
    return None


def request_download(page, progress, should_stop=lambda: False, timeout=120, source_label='表C'):
    """Use normal export menus and accept only a real browser Download event."""
    ui = OnlineFormatter(page, lambda ref: None, progress=progress)
    downloads = []
    listener = lambda download: downloads.append(download)
    page.on('download', listener)
    try:
        opened = exported = selected = confirmed = False
        deadline = time.monotonic() + timeout
        while not downloads:
            if should_stop():
                from sync_runner import SyncStopped
                raise SyncStopped()
            if time.monotonic() >= deadline:
                raise ValueError(f'{source_label}下载超时：未收到Excel文件，请检查{source_label}登录状态、下载权限和导出菜单')
            # Explicit format choices take precedence over generic export menus.
            control = None
            expand = False
            if not selected:
                for label in EXCEL_LABELS:
                    control = _control(ui, label)
                    if control is not None:
                        selected = True
                        break
            if control is None and not exported and not selected:
                for label in EXPORT_LABELS:
                    control = _control(ui, label)
                    if control is not None:
                        exported = True
                        expand = label in ('导出为', '下载为', '导出', '另存为')
                        break
            if control is None and not opened and not exported and not selected:
                control = _main_menu(ui)
                if control is not None:
                    opened = True
                    label = '文件菜单'
            if control is None and (selected or exported) and not confirmed:
                # Never click a generic confirmation outside an export dialog.
                dialogs = ui._visible_dialogs()
                if len(dialogs) == 1:
                    for label in ('下载', '导出', '确定'):
                        control = _control(ui, label, scope=dialogs[0])
                        if control is not None:
                            confirmed = True
                            break
            if control is not None:
                progress(f'正在下载{source_label}：' + label + '…')
                if expand:
                    _expand_export(ui, control, label)
                else:
                    ui._click(control, label)
            else:
                page.wait_for_timeout(250)
        return downloads[0]
    finally:
        page.remove_listener('download', listener)


def save_download(download, root, sheet_name, *, target_name='表C.xlsx', label='表C'):
    """Validate the downloaded workbook before atomically replacing the target file."""
    from format_reference import load
    folder = Path(root) / 'DATA'
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / target_name
    fd, filename = tempfile.mkstemp(prefix='.table_download_', suffix='.xlsx', dir=folder)
    os.close(fd)
    temporary = Path(filename)
    try:
        failure = download.failure()
        if failure:
            raise ValueError(f'{label}下载失败：' + str(failure))
        if Path(download.suggested_filename).suffix.lower() != '.xlsx':
            raise ValueError(f'{label}下载结果不是.xlsx文件，请在{label}导出菜单选择Excel格式')
        download.save_as(str(temporary))
        try:
            reference = load(root, {'source': {'format_file': str(temporary.resolve()),
                                              'format_sheet': sheet_name}})
        except Exception as exc:
            raise ValueError(f'下载的{label}不是有效的Excel工作簿或缺少指定工作表，原{label}保持不变：' + str(exc)) from exc
        try:
            os.replace(temporary, target)
        except PermissionError as exc:
            raise ValueError(f'无法替换 DATA/{target_name}，请关闭正在打开{label}的Excel/WPS后重试') from exc
        return dict(path=f'DATA/{target_name}', sheet=reference['sheet'],
                    size=target.stat().st_size,
                    message=f'{label}已下载到程序根目录 DATA/{target_name}。')
    finally:
        temporary.unlink(missing_ok=True)


def download_table_b(copier, config, page=None):
    """下载表A（线上目标表）为 DATA/表B.xlsx；下载方法完全照搬表C。

    表B 是表A 的原始 xlsx 快照（数据 + 格式），用于在本地与表C 对比找出差异单元格，
    再只把差异粘贴回线上表A。
    """
    from sync_runner import _open_sheet, _stop_requested
    from local_mirror import mirror_path
    if mirror_path(copier.root, config).resolve() == (Path(copier.root) / 'DATA' / '表B.xlsx').resolve():
        raise ValueError('表A1路径不能使用 DATA/表B.xlsx，请为镜像选择另一个文件')
    if page is None:
        page = _open_sheet(copier, config, 'a')
    progress = lambda message: copier.update_upload(message=message)
    progress('正在从表A下载原始Excel文件为表B…')
    download = request_download(page, progress, lambda: _stop_requested(copier), source_label='表A')
    progress('正在保存并校验 DATA/表B.xlsx…')
    return save_download(download, copier.root, config['upload']['sheet_name'],
                         target_name='表B.xlsx', label='表B')


def download_table_c(copier, config, page=None):
    from sync_runner import _open_sheet, _stop_requested
    from local_mirror import mirror_path
    if mirror_path(copier.root, config).resolve() == (Path(copier.root) / 'DATA' / '表C.xlsx').resolve():
        raise ValueError('表A1路径不能使用 DATA/表C.xlsx，请为镜像选择另一个文件')
    if page is None:
        page = _open_sheet(copier, config, 'c')
    progress = lambda message: copier.update_upload(message=message)
    progress('正在从表C下载原始Excel文件为表C…')
    download = request_download(page, progress, lambda: _stop_requested(copier), source_label='表C')
    progress('正在保存并校验 DATA/表C.xlsx…')
    return save_download(download, copier.root, config['source']['sheet_name'],
                         target_name='表C.xlsx', label='表C')
