"""Select an A1 range through the rendered name box, without using clipboard.

Keep an ElementHandle to the inspected input: input.nth(index) can refer to a
different field after the spreadsheet redraws its toolbar. All writes below
are ordinary clicks and keyboard input, never DOM value/readonly overrides.
"""
import re
import time
import urllib.parse

from update_excel import address, column

DETAIL = r'''el => {
 const r=el.getBoundingClientRect(),s=getComputedStyle(el);
 const texts=[el.id,el.className,el.getAttribute('name'),el.title,el.getAttribute('aria-label'),el.placeholder];
 let parents=[];
 for(let p=el.parentElement;p&&parents.length<3;p=p.parentElement){
   parents.push({tag:p.tagName.toLowerCase(),class:String(p.className||'').slice(0,140),role:p.getAttribute('role')||''});
 }
 const hint=texts.filter(Boolean).join(' ')+' '+parents.map(p=>p.class).join(' ');
 const raw=el.value||'',ref=/^(?:\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?[1-9][0-9]*:\$?[1-9][0-9]*)$/i.test(raw.trim())?raw.trim():null;
 const type=(el.type||'text').toLowerCase();
 const explicit=/名称框|名称盒|单元格地址|name[-_\s]?box|cell[-_\s]?(reference|address)|range[-_\s]?(address|box)/i.test(hint);
 const excluded=!!el.closest('[role="dialog"],[aria-modal="true"],[role="gridcell"],td,th') ||
   /password|passwd|search|formula|cell[-_\s]?editor|密码|搜索|公式/i.test(hint)||
   ['username','current-password','new-password','email'].includes(el.autocomplete);
 const shown=el.isConnected&&r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden';
 const x=Math.min(innerWidth-1,Math.max(0,r.left+r.width/2)),y=Math.min(innerHeight-1,Math.max(0,r.top+r.height/2));
 const hit=document.elementFromPoint(x,y),blocked=!!hit&&hit!==el&&!el.contains(hit);
 return {connected:el.isConnected,visible:shown,type,explicit,excluded,
   id:(el.id||'').slice(0,140),class:String(el.className||'').slice(0,180),
   aria_label:(el.getAttribute('aria-label')||'').slice(0,140),parents,
   disabled:el.disabled,readonly:el.readOnly,aria_readonly:el.getAttribute('aria-readonly'),
   focused:document.activeElement===el&&document.hasFocus(),
   address:excluded||type==='password'?null:ref,value_length:raw.length,
   selection_start:el.selectionStart,selection_end:el.selectionEnd,
   upper_left:r.left>=0&&r.left<400&&r.top>=0&&r.top<Math.min(350,innerHeight*.6),
   blocked,blocker:blocked?{tag:hit.tagName.toLowerCase(),class:String(hit.className||'').slice(0,140),role:hit.getAttribute('role')||''}:null};
}'''


class RangeSelectionError(ValueError):
    def __init__(self, code, message, context):
        super().__init__(message)
        self.code = code
        self.diagnostic_context = context


def _frames(page):
    for frame in page.frames:
        host=urllib.parse.urlparse(frame.url).hostname
        if host is None or host in ('kdocs.cn','wps.cn') or host.endswith(('.kdocs.cn','.wps.cn')):
            yield frame


def _dispose(item):
    try:item.dispose()
    except Exception:pass


def _snapshot(detail):
    # Never put arbitrary input values, credentials or copied cell data in a
    # report. Address-like text is included only for the chosen name box.
    return {k:v for k,v in detail.items() if k not in ('value_length','selection_start','selection_end')}


def find_namebox(page, timeout=3):
    deadline=time.monotonic()+timeout
    last=[]
    while True:
        candidates=[];held=[];last=[]
        try:
            for frame in _frames(page):
                # Handles pin DOM nodes. A later insertion cannot shift an
                # input index and silently redirect keyboard input.
                for item in frame.locator('input:not([type="password"]):not([type="hidden"])').element_handles():
                    held.append(item)
                    try:detail=item.evaluate(DETAIL)
                    except Exception:continue  # A frame/input may be redrawing.
                    if not detail['visible'] or detail['type'] not in ('text','search'):continue
                    if len(last)<25:last.append(_snapshot(detail))
                    if detail['excluded']:continue
                    if detail['explicit'] or (detail['address'] and detail['upper_left'] and detail['type']=='text'):
                        candidates.append((frame,item,detail))
            explicit=[c for c in candidates if c[2]['explicit']]
            if explicit:candidates=explicit
            # 名称框被加载遮罩/工具栏重绘盖住时还不能点击：继续等待，避免 select_range
            # 点到一个被遮挡的节点后立刻失败。
            ready=[c for c in candidates if not c[2].get('blocked')]
            if len(ready)==1:
                chosen=ready[0]
                held.remove(chosen[1])
                return chosen
        finally:
            for item in held:_dispose(item)
        if time.monotonic()>=deadline:
            code='NAMEBOX_AMBIGUOUS' if candidates else 'NAMEBOX_NOT_FOUND'
            raise RangeSelectionError(code,'无法唯一识别单元格地址框，请确认表格工具栏已加载。报告已记录可见输入框的控件状态。',
                                      {'candidate_count':len(candidates),'controls':last})
        page.wait_for_timeout(150)


def _normalize(value):
    return (value or '').replace('$','').upper().strip()


def selection_reference(ref):
    """Return a validated command and the anchor WPS displays after Enter."""
    ref=_normalize(ref)
    if re.fullmatch(r'[A-Z]{1,3}:[A-Z]{1,3}',ref):
        left,right=ref.split(':')
        if column(right)<column(left):raise ValueError('整列选区终点不能在起点之前')
        return ref,left+'1'
    if re.fullmatch(r'[1-9][0-9]*:[1-9][0-9]*',ref):
        top,bottom=ref.split(':')
        if not 1<=int(top)<=int(bottom)<=1048576:raise ValueError('整行选区无效或超出范围')
        return ref,'A'+top
    parts=ref.split(':')
    if len(parts) not in (1,2):raise ValueError('选区地址无效')
    if len(parts)==2 and parts[0]==parts[1]:
        # WPS名称框不接受“A1:A1”这类单格范围：折叠为单格地址。
        ref=parts[0];parts=[parts[0]]
    a,b=address(parts[0]),address(parts[-1])
    if b[0]<a[0] or b[1]<a[1]:raise ValueError('选区终点不能位于起点左上方')
    return ref,parts[0]


def select_range(page, ref, *, progress=lambda text:None, attempts=2, find_timeout=6):
    ref,anchor=selection_reference(ref)
    last_error=None
    for attempt in range(attempts):
        item=None;stage='查找地址框';detail={}
        try:
            progress(f'正在通过单元格地址框选择 {ref}'+('（重新定位控件）' if attempt else '')+'…')
            frame,item,detail=find_namebox(page,find_timeout)
            stage='点击地址框'
            # 大表加载时工具栏可能仍在重绘：点击失败就重新定位并再试，
            # 而不是把一次瞬时遮挡/重绘当成整次失败。
            for click_try in range(3):
                try:
                    item.click(timeout=5000)
                    break
                except Exception as exc:
                    if type(exc).__name__=='TargetClosedError':raise
                    _dispose(item);item=None
                    if click_try==2:raise
                    page.wait_for_timeout(300)
                    frame,item,detail=find_namebox(page,find_timeout)
            _dispose(item);item=None
            # Focus/click may replace the input. Re-identify the new live node.
            frame,item,detail=find_namebox(page,find_timeout)
            stage='检查输入焦点'
            if not detail['focused']:
                raise RangeSelectionError('NAMEBOX_FOCUS_LOST','地址框点击后未获得输入焦点，已停止，避免把选区地址写入表格。',{})
            if detail['disabled'] or detail['readonly']:
                raise RangeSelectionError('NAMEBOX_NOT_EDITABLE','当前地址框不可输入，请关闭弹窗并等待表格加载；程序未强制修改控件。',{})
            stage='选中地址框原内容'
            page.keyboard.press('Control+a')
            detail=item.evaluate(DETAIL)
            if not detail['connected'] or not detail['focused'] or detail['readonly'] or detail['disabled']:
                raise RangeSelectionError('NAMEBOX_FOCUS_LOST','地址框重绘或焦点变化，尚未输入选区地址。',{})
            if detail['selection_start']!=0 or detail['selection_end']!=detail['value_length']:
                raise RangeSelectionError('NAMEBOX_INPUT_REJECTED','地址框没有选中原内容，已停止，避免拼接错误地址。',{})
            stage='输入选区地址'
            # Native text input does not overwrite the Excel clipboard. It also
            # avoids fill() waiting on inaccurate aria-readonly annotations.
            page.keyboard.insert_text(ref)
            _dispose(item);item=None
            frame,item,detail=find_namebox(page,find_timeout)
            if not detail['focused'] or _normalize(detail['address'])!=ref:
                raise RangeSelectionError('NAMEBOX_INPUT_REJECTED','地址框未完整接收选区地址，已停止；请查看报告中的输入框状态。',{})
            stage='确认选区并等待焦点离开地址框'
            item.press('Enter',timeout=1500)
            _dispose(item);item=None
            for poll in range(12):
                page.wait_for_timeout(150)
                frame,item,detail=find_namebox(page,timeout=0)
                current=_normalize(detail['address'])
                if current in (ref,anchor) and not detail['focused']:
                    # Some WPS skins display only the upper-left address after
                    # Enter. The whole command was checked before submission;
                    # downstream clipboard checks still verify copied extents.
                    return frame
                _dispose(item);item=None
            raise RangeSelectionError('RANGE_NOT_CONFIRMED','地址框提交后未确认选区或焦点仍停在输入框，已停止，尚未执行后续粘贴。',{})
        except Exception as exc:
            if type(exc).__name__=='TargetClosedError':raise
            context={'requested_range':ref,'step':stage,'attempt':attempt+1,'control':_snapshot(detail)}
            if isinstance(exc,RangeSelectionError):
                exc.diagnostic_context.update(context)
                last_error=exc
            else:
                last_error=RangeSelectionError('NAMEBOX_INTERACTION_FAILED',
                    f'单元格地址框操作失败：{stage}（{ref}）。可能存在控件重绘或遮挡，报告已记录控件状态。',context)
                last_error.__cause__=exc
            if attempt+1<attempts:
                try:
                    # WPS可能弹出“请输入有效的引用位置”等提示：先按Escape清除再重试。
                    page.keyboard.press('Escape')
                except Exception:
                    pass
                page.wait_for_timeout(250)
        finally:
            if item is not None:_dispose(item)
    raise last_error
