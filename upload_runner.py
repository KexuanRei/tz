"""Upload A1 values with C formatting, verify batches and set online dimensions."""
from pathlib import Path
import time
import urllib.parse

from copy_range import column_name
from update_excel import address
from read_limits import check_cells
from file_source import source_path
from format_reference import load as load_reference
from upload_layout import plan_layout, plan_batches
from upload_verify import snapshot, mismatch, differences, merge_ref, expected_merges
from upload_target import plan_upload


# 同一处不一致必须持续这么久才判定失败：大块粘贴后WPS是分批渲染的，
# 偶尔某个单元格要过几秒才写入，过早判定就会误报“回读校验失败”。
MISMATCH_SETTLE_SECONDS = 8.0
# 单批读回的总上限；到点仍不一致就报错，避免无限等待。
READBACK_TIMEOUT_SECONDS = 30.0
# 网页编辑器单批的粘贴上限（按单元格数）：一次往表C贴太多格会让标签页直接崩掉
# （表现为 TargetClosedError / 浏览器已关闭）。实测每批1000行×80列=8万格稳定，
# 5000行×80列=40万格必定崩溃，这里按8万格封顶，超出自动按列数折算每批行数。
ONLINE_PASTE_CELLS = 80_000


def online_batch_rows(requested, columns):
    """把“每批行数”限制在网页编辑器承受得住的范围内（按单元格数折算）。"""
    try:
        requested = int(requested)
    except (TypeError, ValueError):
        requested = 1000
    requested = max(1, requested)
    columns = max(1, int(columns))
    return max(1, min(requested, ONLINE_PASTE_CELLS // columns))


def _has_values(rows):
    return any(str(cell).strip() for row in rows for cell in row)


def _formula_like_cells(expected, first_col, first_row, width):
    """Text cells that start with '=': WPS evaluates them on paste, so they need
    an explicit string formula ("=...") written into the cell afterwards."""
    found = []
    for y, row in enumerate(expected.get('rows') or ()):
        for x in range(min(width, len(row))):
            value = row[x]
            if isinstance(value, str) and value.lstrip().startswith('='):
                found.append((f'{column_name(first_col+x)}{first_row+y}', value))
    return found


def _write_text_cells(copier, page, cells):
    """Write literal text that begins with '=' so WPS keeps it as text.

    A cell typed as "=E1&I1" becomes a formula and evaluates to an empty string in
    table C; entering it as ="=E1&I1" stores exactly the original text instead."""
    for ref, text in cells:
        literal = '="' + str(text).replace('"', '""') + '"'
        copier._select(page, ref)
        page.keyboard.insert_text(literal)
        page.keyboard.press('Enter')
        page.wait_for_timeout(120)


def _hidden_tail(found, actual, origin):
    """差异是否全是“本批有内容、表C读回为空”，并且从第一处差异起往后整片为空。

    这是表C该段行被隐藏/筛选的典型表现：WPS 复制隐藏行时不返回单元格内容，
    于是从隐藏区起点开始，后面整片都读成空白。前半批能核对，尾部无法核对。"""
    if not found:
        return False
    first_row = None
    for item in found:
        if item.get('kind') != 'value':
            return False
        if str(item.get('actual')).strip() or not str(item.get('expected')).strip():
            return False
        try:
            _column, row = address(item['cell'])
        except (KeyError, TypeError, ValueError):
            return False
        first_row = row if first_row is None else min(first_row, row)
    if first_row is None:
        return False
    rows = actual.get('rows') or []
    start = first_row - address(origin)[1]
    if start <= 0 or start >= len(rows):
        return False
    return not _has_values(rows[start:])


def _read_back(copier, page, ref, expected, max_cells, *, settle=None, timeout=None,
               clock=time.monotonic, progress=None):
    """Read again while WPS finishes a paste, without re-pasting the batch.

    大块粘贴后WPS会异步渲染：只要页面还在改善（不一致的单元格变少）就继续读回，
    只有同一处不一致连续稳定settle秒、期间没有任何改善，才判定失败；
    最长等待timeout秒。这样既不会把“还没粘贴完”误判为错误，也不会无限等待。"""
    settle = MISMATCH_SETTLE_SECONDS if settle is None else settle
    timeout = READBACK_TIMEOUT_SECONDS if timeout is None else timeout
    origin = ref.split(':')[0]
    # 纯文本剪贴板少了几行时，尾部按空白补齐并标记未核对：只比较文本描述过的行。
    keep_rows = None
    if int(expected.get('unverified_rows') or 0):
        candidate = expected['row_count'] - int(expected['unverified_rows'])
        if candidate > 0:
            keep_rows = candidate
    start = clock()
    last = None
    first_seen = None
    samples = 0
    best = None
    while True:
        now = clock()
        try:
            actual = snapshot(copier._capture(page, ref), ref, max_cells)
            wanted, seen = expected, actual
            if keep_rows and actual['row_count'] >= keep_rows:
                wanted = dict(expected, rows=expected['rows'][:keep_rows], row_count=keep_rows)
                seen = dict(actual, rows=actual['rows'][:keep_rows], row_count=keep_rows)
            found = differences(wanted, seen, origin)
            diff = mismatch(wanted, seen, origin) if found else None
        except ValueError as exc:
            actual = None
            found = []
            diff = str(exc).replace('表A1未修改', '本批尚未验证通过')
        if actual is not None and not found:
            return actual
        if actual is not None and found and not _has_values(actual['rows']) \
                and _has_values(expected['rows']):
            # 表C这些行被隐藏或被筛选时，WPS复制出来的内容整片为空：无法回读核对，
            # 但粘贴本身是整块写入的。标记为“未核对”继续，避免整次上传因隐藏行中断。
            actual['unverified'] = '表C该区域的行被隐藏或被筛选，WPS复制隐藏行时不返回内容'
            return actual
        if actual is not None and _hidden_tail(found, actual, origin):
            # 只有从某一行起往后整片读回为空：同样是表C该段被隐藏/筛选（WPS复制
            # 隐藏行不返回内容），前半批能核对、后半批无法核对，继续上传并标记未核对。
            actual['unverified'] = ('表C从本批中途起有行被隐藏或被筛选，WPS复制隐藏行时不返回内容；'
                                    '这一段的尾部按未核对处理')
            return actual
        count = len(found) if actual is not None else None
        # 只要不一致在减少，就说明粘贴还在继续进行，重新开始计时。
        improved = count is not None and best is not None and count < best
        if diff != last or improved:
            last, first_seen, samples = diff, now, 1
        else:
            samples += 1
        if count is not None:
            best = count if best is None else min(best, count)
        waited = now - start
        stable = now - first_seen if first_seen is not None else 0.0
        if (samples >= 2 and stable >= settle) or waited >= timeout:
            raise ValueError('表C回读校验失败：' + str(diff))
        if progress is not None:
            progress(waited)
        page.wait_for_timeout(min(400 + 400 * samples, 1500))


def _drop_merges(copier, page, formatter, target_ref, origin, max_cells):
    """取消本批范围内的合并，避免粘贴只写入合并区域左上角、其余单元格留空。

    失败重贴前调用；表C要求保留的合并会在本批校验通过后重新合并。"""
    state = snapshot(copier._capture(page, target_ref), target_ref, max_cells)
    removed = 0
    for item in sorted(state['merges']):
        formatter.unmerge(merge_ref(item, origin))
        removed += 1
    return removed


def _reopen_if_closed(copier, page, formatter, rebuild, upload):
    """表C标签页被关掉/崩溃时重新打开；返回可用的页面与格式器。

    WPS网页版在一次粘贴的行数太多时会直接把标签页搞崩（浏览器日志里是
    TargetClosedError）。崩在批次之间时没必要让整次上传失败：重新打开表C、
    重新定位工作表后继续下一批，已粘贴的批次不受影响。"""
    try:
        # is_closed() 只在明确返回 True 时才算关闭（测试替身可能给出别的东西）。
        if page is not None and page.is_closed() is not True:
            page.title()  # 真正碰一下页面：崩溃的标签页在这里就会抛错
            return page, formatter
    except Exception:
        pass
    copier.update_upload(message='表C页面已关闭或已崩溃，正在重新打开表C并继续上传…')
    page = copier._page_c(upload['share_url'], reload=True)
    page.bring_to_front()
    formatter = rebuild(page)
    formatter.select_sheet(upload['sheet_name'])
    return page, formatter


def _paste_batch(copier, page, formatter, index, total, top, target_ref, expected, sequence,
                 max_cells, on_paste=None, restore_clipboard=None, repair=None):
    """Paste one batch and verify it; on a mismatch repair merges and paste once more.

    WPS粘贴到已有合并区域时只会写入左上角，其他单元格保持旧值，回读就会只差几个
    单元格。这里先取消本批范围内的合并再重贴一次，避免整次上传因个别单元格失败。"""
    import win32clipboard
    from retry_com import retry_call
    attempts = 0
    while True:
        attempts += 1
        if attempts > 1 and restore_clipboard is not None:
            # 读回操作会用WPS选区内容覆盖剪贴板，重贴前必须重新复制表A1本批内容。
            copier.update_upload(message=f'第{index}批正在重新复制表A1本批内容…')
            sequence = restore_clipboard()
        copier._select(page, top)
        current = retry_call(win32clipboard.GetClipboardSequenceNumber)
        if current != sequence:
            raise ValueError(f'第{index}批剪贴板被其他程序改动，已停止上传')
        page.keyboard.press('Control+v')
        if on_paste is not None:
            on_paste()
        page.wait_for_timeout(500)
        if repair is not None:
            # 以“=”开头的文本会被WPS当公式算成空值：粘贴后用字符串公式补写。
            repair()
        try:
            return _read_back(
                copier, page, target_ref, expected, max_cells,
                progress=lambda seconds: copier.update_upload(
                    message=f'第{index}/{total}批已粘贴，正在等表C完成渲染（已等待{seconds:.0f}秒）…'))
        except ValueError as exc:
            if attempts >= 2:
                raise ValueError(
                    f'{exc}；已自动取消本批范围内的合并并重贴1次，仍不一致。'
                    '请检查该单元格在表C是否为合并单元格、受保护或被格式锁定，'
                    '或先清空表C目标区域后再重新上传') from exc
            copier.update_upload(message=f'第{index}批读回不一致，正在取消本批范围内的合并并重新粘贴同一批一次'
                                         f'（{str(exc)[:60]}）…')
            try:
                _drop_merges(copier, page, formatter, target_ref, top, max_cells)
            except ValueError as merge_exc:
                copier.update_upload(message=f'第{index}批合并清理未完成（{str(merge_exc)[:60]}），仍继续重贴…')


def _clear_destination(copier, config, upload, area):
    """空白覆盖：把表C目标区域整块清空。

    “源区域没有数据时按空白覆盖”用的就是这条路径：表A1那一段是空白，相当于
    把表C同位置的内容全部覆盖成空白（=清空）。执行前由用户显式勾选才会走到这里。"""
    from online_format import OnlineFormatter
    from clipboard_html import table_fragment, CopyIntegrityError
    from excel_clipboard import _TableShape, _check_html_errors
    from lxml import etree
    target_ref = area['target_range']
    copier.update_upload(message=f'“空白覆盖”已启用：正在清空表C {target_ref}（该区域原有内容会被删除）…')
    page = copier._page_c(upload['share_url'], reload=True)
    page.bring_to_front()
    formatter = OnlineFormatter(page, lambda ref: copier._select(page, ref),
                                progress=lambda message: copier.update_upload(message=message),
                                diagnostic_dir=copier.root / 'error_reports')
    formatter.select_sheet(upload['sheet_name'])
    copier._select(page, target_ref)
    page.keyboard.press('Delete')
    page.wait_for_timeout(800)
    probe_rows = min(area['rows'], 50)
    col, row = address(area['destination'])
    probe = f'{column_name(col)}{row}:{column_name(col+area["columns"]-1)}{row+probe_rows-1}'
    # 抽查：读回那一段。注意表C的列数可能比请求的少（例如表C只有80列，请求82列），
    # 网页会按实际存在的列回读——空白检查只看“还有没有内容”，不要求列数完全一致。
    clip = copier._capture(page, probe)
    if (clip.get('text') or '').strip():
        raise ValueError(f'“空白覆盖”已执行，但表C {probe} 仍有内容：请检查该区域是否有合并单元格、'
                         f'工作表保护或被筛选；表C其余部分未改动')
    html = clip.get('html') or ''
    if html:
        shape = None
        try:
            shape = _TableShape(config['source']['max_cells'])
            parser = etree.HTMLParser(target=shape, no_network=True)
            fragment = table_fragment(html)
            for offset in range(0, len(fragment), 65536):
                parser.feed(fragment[offset:offset + 65536])
                _check_html_errors(parser)
            parser.close()
            _check_html_errors(parser)
        except (CopyIntegrityError, ValueError):
            raise
        except Exception:
            shape = None  # 表格解析不了时，纯文本已为空，按通过处理
        if shape is not None and (shape.last_text_row >= 0 or shape.last_text_col >= 0):
            raise ValueError(f'“空白覆盖”已执行，但表C {probe} 仍有内容：请检查该区域是否有合并单元格、'
                             f'工作表保护或被筛选；表C其余部分未改动')
    return dict(rows=area['rows'], columns=area['columns'], destination=area['destination'],
                requested_destination=area['requested_destination'], target_range=target_ref,
                source_range=area['source_range'], truncated=False, verified_rows=probe_rows,
                format_source='空白覆盖', merges=0, dimensions=None, freeze_applied=False,
                unverified_batches=0, partial_batches=0, blank_overwrite=True, batches=0)


def run_upload(copier, config, destination):
    from openpyxl import load_workbook
    from upload_format import prepared_upload
    from online_format import OnlineFormatter
    from excel_clipboard import read_clipboard
    import pythoncom
    import win32com.client

    upload = config['upload']
    copier.update_upload(message='正在检查表A1路径并读取表C格式…')
    a1 = source_path(copier.root, config)
    reference = load_reference(copier.root, config)
    if reference is None:
        raise ValueError('请先填写表C格式文件并保存配置')
    source_range = (upload.get('source_range') or '').strip().upper() or None
    format_path = (Path(copier.root) / config['source']['format_file']).resolve()
    book = load_workbook(a1, read_only=True)
    try:
        if config['source']['sheet_name'] not in book.sheetnames:
            raise ValueError('表A1中找不到所配置的工作表')
        sheet = book[config['source']['sheet_name']]
        max_row, max_col = sheet.max_row or 1, sheet.max_column or 1
        # 表A1末尾常留着一大片空白行（例如数据到47763行、表却到50000行）：
        # 先找出最后一行有内容的行，不再把这些空白行往表C里粘贴。
        data_rows = {}  # 行号 -> (最左有内容的列, 最右有内容的列)
        for index, values in enumerate(sheet.iter_rows(min_row=1, max_row=max_row,
                                                       max_col=max_col, values_only=True), start=1):
            left = right = None
            for position, value in enumerate(values, start=1):
                if value is not None and value != '':
                    if left is None:
                        left = position
                    right = position
            if left is not None:
                data_rows[index] = (left, right)
        last_data = max(data_rows) if data_rows else 0
        area = plan_upload(last_data or max_row, max_col, destination,
                           upload.get('overflow', 'stop'), source_range=source_range,
                           allow_blank=bool(upload.get('blank_overwrite')))
        if area.get('blank_overwrite'):
            return _clear_destination(copier, config, upload, area)
        rows, cols = area['rows'], area['columns']
        check_cells(rows, cols, config['source']['max_cells'], '表A1上传范围')
        src_left, src_top = area['source_left'], area['source_top']
        has_data = any(src_left <= right and left <= src_left + cols - 1
                       for row, (left, right) in data_rows.items()
                       if src_top <= row <= src_top + rows - 1)
        trimmed_tail = max_row - last_data if last_data else 0
    finally:
        book.close()
    if not rows or not cols:
        raise ValueError('表A1本次上传范围为空，请检查源区域或目标区域大小')
    if not has_data:
        # 个别空白单元格/空白行照常上传；只有整个范围都没有任何内容时才停止，
        # 否则一次失败的复制会让表C整片被空白覆盖。
        raise ValueError('表A1本次上传范围没有数据（全部为空白单元格），已停止以免清空表C；'
                         '请确认表A1是否已完成复制，或缩小源区域后重试')
    if trimmed_tail > 0:
        copier.update_upload(message=f'表A1第{last_data + 1}行起全是空白单元格，已按实际数据范围'
                                     f'（到第{last_data}行）上传，不再粘贴末尾空白行。')
    max_cells = config['source']['max_cells']
    check_cells(rows, cols, max_cells, '表A1上传范围')
    source_ref = area['source_range']
    destination = area['destination']
    # Bounds/merge preflight occurs before any browser write.
    plan_layout(reference, source_ref, destination)
    # The scratch copy is 1-based; table C merges are translated to it so batch
    # boundaries never cut a merge even when the source region is offset.
    internal_merges = [(m[0]-src_left+1, m[1]-src_top+1, m[2]-src_left+1, m[3]-src_top+1)
                       for m in reference.get('merges', ())
                       if m[2] >= src_left and m[0] < src_left+cols and m[3] >= src_top and m[1] < src_top+rows]
    batch_rows=online_batch_rows(upload['batch_rows'], cols)
    batches = plan_batches((1,1), rows, cols, batch_rows,
                           merges=internal_merges, max_cells=min(max_cells,1_000_000))
    try:
        requested_rows=int(upload['batch_rows'])
    except (TypeError,ValueError):
        requested_rows=batch_rows
    # 每批行数填太大时自动折算，并把折算结果留在准备阶段的状态栏里，避免用户以为设置没生效。
    clamp_note=''
    if requested_rows>batch_rows:
        clamp_note=(f'；每批{requested_rows:,}行 × {cols:,}列＝{requested_rows*cols:,}格太大'
                    f'（网页编辑器会崩），已按每批{batch_rows:,}行上传')
    dest_col, dest_row = address(destination)
    copier.update_upload(message=f'正在准备上传副本：表A1的{source_ref} → 表C的{area["target_range"]}；格式重新读取表C…'
                         + ('已按目标区域大小截取数据。' if area['truncated'] else '') + clamp_note)
    pythoncom.CoInitialize()
    excel = None
    pasted = []  # 已真正按下粘贴的批次序号，用于失败时说明可能保留了哪些批次
    unverified = []  # 因表C隐藏行/筛选无法回读核对的批次
    partial = []  # 纯文本剪贴板少行、尾部按空白补齐的批次
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        excel.AskToUpdateLinks = False
        with prepared_upload(excel, a1, config['source']['sheet_name'], format_path,
                             config['source'].get('format_sheet',''), rows, cols, reference,
                             source_left=src_left, source_top=src_top) as prepared:
            sheetx = prepared.sheet
            # prepared.layout uses 1-based scratch coordinates; map it from A1.
            plan = plan_layout(prepared.layout, f'A1:{column_name(cols)}{rows}', destination)
            # The user may have navigated this tab elsewhere since preparation.
            # Always navigate to the configured destination before any write.
            copier.update_upload(message='上传副本已准备，正在打开并等待表C网页加载…')
            page = copier._page_c(upload['share_url'],reload=True)
            page.bring_to_front()
            url = urllib.parse.urlparse(page.url)
            if url.scheme != 'https' or url.hostname not in {'www.kdocs.cn','kdocs.cn'}:
                raise ValueError('请先点击“打开表C并登录”，完成登录后再上传')
            copier.context.grant_permissions(['clipboard-read','clipboard-write'],origin=url.scheme+'://'+url.netloc)

            def make_formatter(page):
                return OnlineFormatter(page, lambda ref_:copier._select(page,ref_),
                                       progress=lambda message:copier.update_upload(message=message),
                                       diagnostic_dir=copier.root/'error_reports')
            formatter = make_formatter(page)
            copier.update_upload(message='正在定位表C工作表：'+upload['sheet_name'])
            formatter.select_sheet(upload['sheet_name'])
            if upload.get('skip_dimensions'):
                copier.update_upload(message='已选择跳过表C尺寸设置：行高、列宽请手动设置，数据与合并仍自动上传核对…')
            else:
                copier.update_upload(message='正在检查表C的行高、列宽菜单及尺寸单位…')
                formatter.preflight(plan)
            for index,(ref,row,count) in enumerate(batches,1):
                # 表C标签页被关掉或崩掉时（一次贴太多行时最常见）先重新打开，
                # 否则整次上传会以 TargetClosedError 中断在半路。
                page, formatter = _reopen_if_closed(copier, page, formatter, make_formatter, upload)
                top = f'{column_name(dest_col)}{dest_row+row-1}'
                bottom = f'{column_name(dest_col+cols-1)}{dest_row+row+count-2}'
                target_ref = f'{top}:{bottom}'
                copier.update_upload(message=f'正在上传并校验第{index}/{len(batches)}批：表A1 {ref} → 表C {target_ref}；格式来自表C。')
                if area['bounded']:
                    # Read before putting Excel content onto the clipboard. A
                    # merge-expanded or ambiguous target must not be pasted over.
                    try:
                        existing = snapshot(copier._capture(page, target_ref), target_ref, max_cells)
                        if any(c1 < 0 or r1 < 0 or c2 >= cols or r2 >= count
                               for c1,r1,c2,r2 in existing['merges']):
                            raise ValueError('表C合并区域跨越本批写入边界')
                    except ValueError as exc:
                        from upload_target import UploadTargetError
                        raise UploadTargetError('UPLOAD_TARGET_UNVERIFIED',
                            f'无法确认表C目标区域{target_ref}完整且没有跨边界合并；本批尚未粘贴。请检查合并状态及选区。',
                            requested_destination=area['requested_destination'], target_range=target_ref) from exc
                # Excel忙碌时会拒绝调用、剪贴板也可能被别的程序短暂占用：这些都要重试，
                # 否则传到一半就会以“请检查Excel…”这种笼统提示中断。
                from retry_com import retry_call
                retry_call(lambda: sheetx.Range(ref).Copy())
                clip = retry_call(read_clipboard)
                expected = snapshot(clip,ref,max_cells)
                if expected['merges'] != expected_merges({'merges': internal_merges}, row, count, cols):
                    raise ValueError(f'上传副本在{ref}的合并格式与表C不一致，已停止')
                sequence = clip['sequence']
                # 文本单元格以“=”开头时，WPS会当公式算成空值；粘贴后用字符串公式补写，
                # 让表C保存原来的文本，而不是直接中断整次上传。
                formula_text = _formula_like_cells(expected, dest_col, dest_row + row - 1, cols)
                if len(formula_text) > 300:
                    sample = '、'.join(ref for ref, _ in formula_text[:5])
                    raise ValueError(f'本批有{len(formula_text)}个以=开头的文本单元格（{sample} 等），'
                                     '数量过多，请先在表A1里把它们改成实际内容再上传；本批尚未粘贴。')
                repair = None
                if formula_text:
                    copier.update_upload(
                        message=f'第{index}批有{len(formula_text)}个以=开头的文本单元格，'
                                '粘贴后会按文本逐格补写（WPS默认会把它们当公式算成空）。')
                    repair = lambda cells=formula_text: _write_text_cells(copier, page, cells)
                def restore_clipboard(ref=ref, sheet=sheetx):
                    from retry_com import retry_call
                    retry_call(lambda: sheet.Range(ref).Copy())
                    return read_clipboard()['sequence']
                actual = _paste_batch(copier,page,formatter,index,len(batches),top,target_ref,
                                      expected,sequence,max_cells,on_paste=lambda: pasted.append(index),
                                      restore_clipboard=restore_clipboard,repair=repair)
                if expected.get('unverified_rows'):
                    partial.append(index)
                    copier.update_upload(
                        message=f'第{index}批的纯文本剪贴板缺少最后{expected["unverified_rows"]}行'
                                '（末尾空白行），已按空白补入并继续；这几行按未核对处理。')
                if actual.get('unverified'):
                    unverified.append(index)
                    copier.update_upload(
                        message=f'第{index}批读回为空：{actual["unverified"]}。已按整块粘贴继续，'
                                '稍后可取消隐藏/清除筛选后重新上传以复核。')
                elif actual['merges'] != expected['merges']:
                    # Only change merges contained in a successfully read-back
                    # rectangle, whose values already match the source.
                    for item in sorted(actual['merges']-expected['merges']):
                        formatter.unmerge(merge_ref(item,top))
                    for item in sorted(expected['merges']-actual['merges']):
                        formatter.merge(merge_ref(item,top))
                    actual = _read_back(copier,page,target_ref,expected,max_cells)
                    if actual['merges'] != expected['merges']:
                        raise ValueError(f'表C在{target_ref}的合并格式仍与表C不一致，已停止')
            if upload.get('skip_dimensions'):
                copier.update_upload(message='各批数据与合并格式已核对，已跳过表C行高/列宽设置。')
                dimension_stats = None
            else:
                copier.update_upload(message='各批数据与合并格式已核对，正在按表C设置表C的行高、列宽…')
                dimension_stats = formatter.apply_dimensions(plan)
            return dict(rows=rows,columns=cols,destination=destination,batches=len(batches),
                        requested_destination=area['requested_destination'], target_range=area['target_range'],
                        source_range=source_ref,truncated=area['truncated'],
                        verified_rows=rows,format_source='表C',merges=len(plan['merges']),
                        dimensions=dimension_stats,freeze_applied=False,
                        unverified_batches=len(unverified), partial_batches=len(partial))
    except Exception as exc:
        if isinstance(exc, ValueError):
            detail = str(exc)
        elif isinstance(exc, MemoryError):
            detail = '电脑可用内存不足，无法完成上传副本准备；请关闭占用内存的程序或缩小表A1实际范围'
        else:
            # 把真实异常类型和内容带出来，否则只看到“请检查Excel/登录/权限”无法判断。
            if type(exc).__name__ == 'TargetClosedError':
                detail = (f'专用Edge的表C标签页已关闭或崩溃（{str(exc)[:120]}）；'
                          '一次往表C粘贴的行数太多时最常见。程序已自动把每批限制在'
                          f'{ONLINE_PASTE_CELLS:,}格以内，请重新点“打开表C并登录”，'
                          '并把“上传每批行数”调到2000以内再试')
            else:
                detail = (f'{type(exc).__name__}: {str(exc)[:200]}；'
                          '请检查Excel是否被弹窗/对话框占用、表C登录状态和编辑权限')
        written = pasted[-1] if pasted else 0
        phase = f'表C已尝试写入{written}批，可能保留部分更新，请修正后重新上传。' if written else '尚未向表C粘贴数据。'
        raise ValueError(detail+'；'+phase) from exc
    finally:
        if excel is not None:
            try:excel.Quit()
            except Exception:pass
        pythoncom.CoUninitialize()
