"""Apply layout using visible spreadsheet controls, never private WPS APIs.

The adapter deliberately stops for ambiguous controls, unknown measurement
units or unverifiable results.  Synthetic DOM tests exercise the interaction
contract; individual WPS page versions still require live compatibility checks.
"""
import datetime as dt
import math
import re
import time
import urllib.parse

from copy_range import column_name
from update_excel import address


CONTROL_SELECTOR = 'button,[role="button"],[role="menuitem"],[title],[aria-label],[data-tooltip]'
INVENTORY_SELECTOR = ('button,[role="button"],[role="menuitem"],[role="option"],[title],[aria-label],[data-tooltip],'
                      '[onclick],[tabindex]:not([tabindex="-1"])')
DIALOG_SELECTOR = '[role="dialog"],dialog[open],[aria-modal="true"],.ant-modal,.el-dialog,.dialog'
ROW_LABELS = ('行高', '设置行高', '调整行高', 'Row Height')
COLUMN_LABELS = ('列宽', '设置列宽', '调整列宽', 'Column Width')
SIZE_MENUS = ('行和列', '行列', '行与列', 'Rows and Columns', 'Row and Column')
CONFIRM_LABELS = ('确定', '确认', 'OK', 'Confirm')
CANCEL_LABELS = ('取消', '关闭', 'Cancel', 'Close')
SHEET_SELECTOR = ('[role="tab"],[role="tablist"] button,[class*="sheet-tab" i],'
                  '[class*="sheetTab" i],[class*="sheet_tab" i],[class*="sheet-name" i],'
                  '[class*="sheetName" i],[class*="sheet-item" i],[data-sheet-name],'
                  '[data-sheet-id],[class*="sheet-list" i] > *,[class*="sheetlist" i] > *')
SHEET_SCAN = r'''(elements, wanted) => {
 const cls = el=>String(el.className||'').replace(/([a-z])([A-Z])/g,'$1-$2').toLowerCase();
 const visible=el=>el.getClientRects().length && getComputedStyle(el).visibility!=='hidden';
 const bar=el=>el.getAttribute('role')==='tablist'||
   /(?:sheet-names?-(?:wrap|list|container)|et-status-bar-main|sheet-tab-(?:bar|list|container))/.test(cls(el));
 const arc=el=>el.classList.contains('status-sheet-item-selected-arc');
 const wpsName=el=>el.classList.contains('sheet-name');
 const selectionOwner=el=>{
   if(!wpsName(el))return el;
   // This WPS skin puts the selected arc beside the name. Associate it with
   // exactly one label below the shared sheet bar; never borrow another tab's
   // marker, even when that other label is hidden or outside the viewport.
   let boundary=el.parentElement;
   while(boundary&&boundary!==document.body&&!bar(boundary))boundary=boundary.parentElement;
   if(!boundary||boundary===document.body)return el;
   let result=el;
   for(let p=el.parentElement;p&&p!==boundary;p=p.parentElement){
     if(p.querySelectorAll('.sheet-name').length!==1)break;
     result=p;
   }
   // A single-sheet variant puts both nodes directly in the bar without an
   // item wrapper. Accept only a sole name and its direct-sibling arc; adding
   // another name (even hidden) invalidates this association.
   if(result===el&&el.parentElement===boundary&&boundary.querySelectorAll('.sheet-name').length===1&&
      Array.from(boundary.children).some(arc)&&
      !Array.from(boundary.querySelectorAll('[role="tab"],[data-sheet-name]')).some(node=>node!==el))return boundary;
   return result;
 };
 const visibleArcs=root=>Array.from(root.querySelectorAll('.status-sheet-item-selected-arc')).filter(marker=>{
   if(!visible(marker))return false;
   for(let p=marker;p;p=p.parentElement){
     const s=getComputedStyle(p);
     if(s.display==='none'||s.visibility==='hidden'||s.visibility==='collapse'||Number(s.opacity)===0)return false;
     if(p===root)break;
   }
   return true;
 });
 const labels=el=>[el.getAttribute('data-sheet-name'),el.getAttribute('aria-label'),el.title,el.innerText]
   .filter(x=>typeof x==='string'&&x.trim()&&x.trim().length<=100).map(x=>x.trim());
 const owner=el=>el.getAttribute('role')==='tab'||
   (/sheet[-_]?(tab|item|name)/.test(cls(el))&&!bar(el)&&!arc(el)&&!/(container|wrapper|panel|tablist|tab-list|tab-bar)/.test(cls(el)));
 const selected=el=>{
   const scope=selectionOwner(el);
   for(let p=el;p;p=p.parentElement){
     if(p.getAttribute('aria-selected')==='false'||p.getAttribute('data-selected')==='false')return false;
     if(p===scope)break;
   }
   if(['true'].includes(el.getAttribute('aria-selected'))||el.getAttribute('data-selected')==='true')return true;
   if(['active','selected','current'].includes(el.getAttribute('data-state')))return true;
   if(wpsName(el)&&scope!==el&&visibleArcs(scope).length>0)return true;
   return owner(el)&&/(?:^|[\s_-])(active|selected|current)(?:$|[\s_-])/.test(cls(el));
 };
 const path=el=>{let p=[];while(el&&el.nodeType===1){let i=1,s=el.previousElementSibling;
   while(s){if(s.tagName===el.tagName)i++;s=s.previousElementSibling}
   p.unshift(el.tagName.toLowerCase()+':nth-of-type('+i+')');el=el.parentElement}return p.join(' > ')};
 const tabs=elements.filter(el=>visible(el)&&!bar(el)&&!arc(el)&&!el.closest('[role="grid"],[role="gridcell"],td,th')&&
   (owner(el)||el.hasAttribute('data-sheet-name')||el.hasAttribute('data-sheet-id')||el.closest('[role="tablist"]')));
 const matching=tabs.filter(el=>labels(el).includes(wanted));
 // An outer tab and its title span are one control, not two matching sheets.
 const groups=[];
 for(const el of matching){let group=groups.find(g=>g.some(other=>other.contains(el)||el.contains(other)));
   if(group)group.push(el);else groups.push([el])}
 const rank=el=>(el.getAttribute('role')==='tab'?100:0)+(selected(el)?40:0)+(owner(el)?10:0)-
   (/(label|title|text|name)/.test(cls(el))?5:0);
 const describe=el=>{
   const scope=selectionOwner(el),parents=[];
   for(let p=el.parentElement;p&&p!==document.body&&parents.length<5;p=p.parentElement){
     parents.push({tag:p.tagName.toLowerCase(),class:cls(p).slice(0,180),role:p.getAttribute('role')||'',
       aria_selected:p.getAttribute('aria-selected'),shared_bar:bar(p)});
     if(bar(p))break;
   }
   return {path:path(el),labels:labels(el),role:el.getAttribute('role')||'',
     class:cls(el).slice(0,180),selected:selected(el),
     selection_owner_class:cls(scope).slice(0,180),
     visible_selected_arcs:wpsName(el)&&scope!==el?visibleArcs(scope).length:0,parents};
 };
 return {matches:groups.map(g=>describe(g.sort((a,b)=>rank(b)-rank(a))[0])),
   visible_tabs:tabs.slice(0,30).map(describe)};
}'''


class SheetTabError(ValueError):
    def __init__(self, code, message, details):
        super().__init__(message)
        self.code = code
        self.diagnostic_context = details


class DimensionControlError(ValueError):
    code = 'DIMENSION_CONTROL_UNAVAILABLE'

    def __init__(self, message, details, *, code=None):
        super().__init__(message)
        self.diagnostic_context = details
        if code is not None:self.code=code


def _label(text):
    text = re.sub(r'\s+', ' ', text or '').strip()
    text = re.sub(r'\s*\([A-Za-z]\)\s*$', '', text)
    return re.sub(r'(?:\.\.\.|…|：|:)\s*$', '', text).strip().casefold()


def _units(text):
    found = set()
    if re.search(r'像素|\bpixels?\b|\bpx\b', text or '', re.I):
        found.add('px')
    if re.search(r'磅|\bpoints?\b|\bpt\b', text or '', re.I):
        found.add('pt')
    if re.search(r'字符|字元|\bcharacters?\b|\bchars?\b', text or '', re.I):
        found.add('excel_char')
    return found


def _numeric(text):
    try:
        value = float(str(text).strip())
    except (TypeError, ValueError, OverflowError) as exc:
        raise ValueError('表C尺寸对话框中的数值无法识别，已停止') from exc
    if not math.isfinite(value) or value < 0:
        raise ValueError('表C尺寸对话框中的数值无效，已停止')
    return value


class OnlineFormatter:
    def __init__(self, page, select_range, progress=lambda message: None, diagnostic_dir=None):
        self.page = page
        self.select_range = select_range
        self.progress = progress
        self.preflight_units = {}
        self.sheet_timeout = 20
        self.last_ref = None
        self.diagnostic_dir = diagnostic_dir

    def _select(self, ref):
        # 记录最近选区，右键表头兜底时需要知道行号/列标。
        self.last_ref = ref
        return self.select_range(ref)

    def _frames(self):
        # Include about:blank for deterministic local DOM fixtures; on live
        # pages, only the document's kdocs frames are formatting candidates.
        for frame in self.page.frames:
            host = urllib.parse.urlparse(frame.url).hostname
            if host is None or host in ('kdocs.cn','wps.cn') or host.endswith(('.kdocs.cn','.wps.cn')):
                yield frame

    def _pause(self):
        self.page.wait_for_timeout(100)

    def _controls(self, labels, scope=None, *, role=None, menu_only=False):
        wanted = {_label(value) for value in labels}
        roots = [(None, scope)] if scope is not None else [(frame, frame) for frame in self._frames()]
        found = []
        seen = set()
        for frame, root in roots:
            locator = root.locator(CONTROL_SELECTOR)
            # Large sheets may expose thousands of title/aria-label elements.
            # Filter publicly rendered control labels in one browser round trip,
            # before inspecting or clicking the few matching controls.
            matches = locator.evaluate_all('''(elements, wanted) => {
                const normalize = text => (text||'').replace(/\\s+/g,' ').trim()
                    .replace(/\\s*\\([A-Za-z]\\)\\s*$/,'').replace(/(?:\\.\\.\\.|…|：|:)\\s*$/,'').trim().toLowerCase();
                return elements.flatMap((el,index) => {
                    const control = el.closest('button,[role="button"],[role="menuitem"]') || el;
                    const labels = [el.title,el.getAttribute('aria-label'),el.getAttribute('data-tooltip')];
                    if (el.tagName === 'BUTTON' || ['button','menuitem'].includes(el.getAttribute('role')))
                        labels.push(el.innerText);
                    if(!labels.some(label=>wanted.includes(normalize(label))) || !el.getClientRects().length)
                        return [];
                    const style=getComputedStyle(el);
                    if(style.visibility==='hidden'||style.display==='none')return [];
                    let path = [], current = control;
                    while (current && current.nodeType === 1) {
                        let index=1, prev=current.previousElementSibling;
                        while(prev){if(prev.tagName===current.tagName)index++;prev=prev.previousElementSibling;}
                        path.unshift(current.tagName.toLowerCase()+':nth-of-type('+index+')');
                        current=current.parentElement;
                    }
                    return [{index, path:path.join(' > '), role:control.getAttribute('role'),
                            popup:control.getAttribute('aria-haspopup'),
                            disabled:control.disabled || control.getAttribute('aria-disabled')==='true'}];
                });
            }''',list(wanted))
            for info in matches:
                item=locator.nth(info['index'])
                if not item.is_visible():
                    continue
                if role and info['role'] != role:
                    continue
                if menu_only and info['popup'] not in ('true', 'menu'):
                    continue
                key = (id(frame) if frame is not None else id(root), info['path'])
                if key in seen:
                    continue
                seen.add(key)
                # Keep the original locator (a title-bearing child can safely
                # trigger its button) but deduplicate the owning control.
                found.append((item, info))
        return found

    def _control_inventory(self):
        """失败时列出可见交互元素标签（不含单元格内容），用于页面适配。"""
        items = []
        for frame in self._frames():
            try:
                info = frame.locator(INVENTORY_SELECTOR).evaluate_all('''(elements) => {
                    const out=[]; const seen=new Set();
                    for(let i=0;i<elements.length && out.length<60;i++){
                        const el=elements[i];
                        if(el.closest('[role="grid"],[role="gridcell"],td,th'))continue;
                        const r=el.getBoundingClientRect();
                        if(!r.width||!r.height)continue;
                        const control=el.closest('button,[role="button"],[role="menuitem"]')||el;
                        const labels=[el.title,el.getAttribute('aria-label'),el.getAttribute('data-tooltip')];
                        const interactive=el.tagName==='BUTTON'||['button','menuitem','option'].includes(el.getAttribute('role'))
                            ||el.hasAttribute('onclick')||el.hasAttribute('tabindex');
                        if(interactive)labels.push(el.innerText);
                        const text=labels.filter(Boolean).join(' ').replace(/\\s+/g,' ').trim().slice(0,60);
                        if(!text)continue;
                        const key=text+'|'+(control.getAttribute('role')||'');
                        if(seen.has(key))continue; seen.add(key);
                        out.push({tag:control.tagName,role:control.getAttribute('role'),label:text});
                    }
                    return out;
                }''')
                items.extend(info)
            except Exception:
                continue
        return items[:60]

    def _capture_ui_dump(self):
        """尺寸入口全部失败时保存界面截图，便于针对真实页面适配。

        截图只写入本机 error_reports 目录（可能包含表格数据），由用户决定是否反馈。"""
        if self.diagnostic_dir is None:
            return
        try:
            from pathlib import Path
            folder = Path(self.diagnostic_dir)
            folder.mkdir(exist_ok=True)
            stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            path=folder / f'表C界面_{stamp}.png'
            self.page.screenshot(path=str(path))
            return str(path)
        except Exception:
            pass

    def _unique_control(self, labels, scope=None, *, optional=False, role=None, menu_only=False):
        found = self._controls(labels, scope, role=role, menu_only=menu_only)
        if not found and optional:
            return None
        if len(found) != 1:
            exc = ValueError('无法唯一定位表C的“%s”控件；当前网页界面需适配后才能继续' % labels[0])
            exc.diagnostic_context = {'wanted': list(labels), 'matched_controls': len(found),
                                      'visible_controls': self._control_inventory()}
            raise exc
        item, info = found[0]
        if info['disabled'] or not item.is_enabled():
            exc = ValueError('表C的“%s”控件不可用，请确认拥有编辑权限' % labels[0])
            exc.diagnostic_context = {'wanted': list(labels), 'visible_controls': self._control_inventory()}
            raise exc
        return item

    def _header_target(self, wanted, kind):
        """定位行号/列标表头元素，返回Locator（Playwright自动处理iframe坐标）。"""
        for frame in self._frames():
            try:
                matches = frame.locator('div,span').evaluate_all('''(elements, wanted) => {
                    const out=[];
                    for(let i=0;i<elements.length;i++){
                        const el=elements[i];
                        if(el.childElementCount!==0)continue;
                        if((el.textContent||'').trim()!==wanted)continue;
                        const r=el.getBoundingClientRect();
                        if(r.width<6||r.height<6)continue;
                        if(r.width>120||r.height>40)continue;  // 表头单元格是小元素
                        out.push({index:i, x:r.left+r.width/2, y:r.top+r.height/2});
                    }
                    return out;
                }''', wanted)
            except Exception:
                continue
            for info in matches:
                # 列标在表格顶部条带、行号在左侧条带；数据单元格不会出现在这些位置。
                if kind == 'column' and 0 < info['y'] < 250:
                    return frame.locator('div,span').nth(info['index'])
                if kind == 'row' and 0 < info['x'] < 250:
                    return frame.locator('div,span').nth(info['index'])
        return None

    def _toolbar_size_item(self, kind, labels):
        """快捷工具/常规/样式等抽屉里可能藏着行高/列宽入口。"""
        for name in ('快捷工具', '常规', '样式'):
            menu = self._unique_control((name,), optional=True)
            if menu is None:
                continue
            self._click(menu, name)
            for _ in range(10):
                control = self._unique_control(labels, optional=True)
                if control is not None:
                    return control
                self._pause()
            # 这个抽屉没有尺寸入口：关闭后试下一个。
            try:
                self.page.keyboard.press('Escape')
            except Exception:
                pass
            self._pause()
        return None

    def _context_size_item(self, kind, labels):
        """部分WPS皮肤没有顶部“行和列”菜单：点击/右键列标行号打开尺寸入口。"""
        if not self.last_ref:
            return None
        parts = self.last_ref.split(':')
        left, top = address(parts[0])
        wanted = column_name(left) if kind == 'column' else str(top)
        target = self._header_target(wanted, kind)
        if target is None:
            return None
        try:
            # Locator.click自行处理iframe偏移与滚动；page.mouse用的是主页面坐标。
            # 先左键点击表头选中整行/整列：部分皮肤只在整行/整列选中后显示尺寸入口。
            target.click(timeout=3000)
            for _ in range(6):
                control = self._unique_control(labels, optional=True)
                if control is not None:
                    return control
                self._pause()
            target.click(button='right', timeout=3000)
        except Exception:
            return None
        for _ in range(10):
            control = self._unique_control(labels, role='menuitem', optional=True)
            if control is not None:
                return control
            self._pause()
        return None

    def _click(self, item, label):
        try:
            item.click(timeout=3000)
        except Exception as exc:
            raise ValueError('无法操作表C的“%s”控件，已停止' % label) from exc
        self._pause()

    def _visible_dialogs(self):
        dialogs = []
        for frame in self._frames():
            for item in frame.locator(DIALOG_SELECTOR).all():
                if item.is_visible():
                    # A nested framework wrapper for the same dialog is not a
                    # second dialog.  Only the outer visible dialog is counted.
                    nested = item.evaluate('''el => {
                        let p=el.parentElement;
                        while(p){if(p.matches('[role="dialog"],dialog[open],[aria-modal="true"],.ant-modal,.el-dialog,.dialog')
                            && p.getClientRects().length)return true;p=p.parentElement;}return false;
                    }''')
                    if not nested:
                        dialogs.append(item)
        return dialogs

    def _size_popup(self):
        """定位WPS自定义尺寸弹层：不匹配标准对话框选择器时，用
        “确定/确认按钮 + 取消/关闭按钮 + 唯一可编辑数值输入框”
        的共同祖先识别弹层容器。"""
        for frame in self._frames():
            try:
                info = frame.evaluate('''() => {
                    const norm=t=>(t||'').replace(/\\s+/g,' ').trim();
                    const visible=el=>{const r=el.getBoundingClientRect();return r.width&&r.height
                        &&getComputedStyle(el).visibility!=='hidden'&&getComputedStyle(el).display!=='none';};
                    const inSet=(t,set)=>set.includes(norm(t));
                    const btns=Array.from(document.querySelectorAll('button,[role="button"]')).filter(visible);
                    const confirm=btns.filter(b=>inSet(b.innerText,['确定','确认','OK','Confirm'])
                        ||inSet(b.title,['确定','确认','OK','Confirm'])
                        ||inSet(b.getAttribute('aria-label'),['确定','确认','OK','Confirm']));
                    const cancel=btns.filter(b=>inSet(b.innerText,['取消','关闭','Cancel','Close'])
                        ||inSet(b.title,['取消','关闭','Cancel','Close'])
                        ||inSet(b.getAttribute('aria-label'),['取消','关闭','Cancel','Close']));
                    if(confirm.length!==1||cancel.length<1)return null;
                    const inputs=Array.from(document.querySelectorAll('input:not([type="hidden"])')).filter(i=>{
                        if(!visible(i)||i.disabled||i.readOnly)return false;
                        return ['text','number','search'].includes((i.type||'text').toLowerCase());
                    });
                    if(inputs.length!==1)return null;
                    const chain=el=>{const out=[];while(el&&el!==document.body&&el!==document.documentElement){out.push(el);el=el.parentElement;}return out;};
                    const a=chain(inputs[0]),b=chain(confirm[0]);
                    let common=null;
                    for(const node of a){if(b.includes(node)){common=node;break;}}
                    if(!common||!common.contains(cancel[0]))return null;
                    const path=el=>{const parts=[];let cur=el;
                        while(cur&&cur.nodeType===1&&cur!==document.body){let i=1,s=cur.previousElementSibling;
                            while(s){if(s.tagName===cur.tagName)i++;s=s.previousElementSibling;}
                            parts.unshift(cur.tagName.toLowerCase()+':nth-of-type('+i+')');cur=cur.parentElement;}
                        return 'body > '+parts.join(' > ');};
                    return {container:path(common), text:common.innerText.slice(0,400)};
                }''')
            except Exception:
                continue
            if not info or not info.get('container'):
                continue
            container = frame.locator(info['container'])
            try:
                if container.count() != 1 or not container.is_visible():
                    continue
            except Exception:
                continue
            return container
        return None

    def _wait_popup(self):
        for _ in range(20):
            if self._visible_dialogs() or self._size_popup() is not None:
                return True
            self._pause()
        return None

    def _dialog(self):
        for _ in range(20):
            dialogs = self._visible_dialogs()
            if len(dialogs) == 1:
                return dialogs[0]
            if len(dialogs) > 1:
                raise ValueError('表C同时存在多个对话框，无法确认尺寸设置对象，已停止')
            # WPS自定义弹层不匹配标准对话框选择器，用按钮+输入框共同祖先识别。
            popup = self._size_popup()
            if popup is not None:
                return popup
            self._pause()
        raise ValueError('表C未出现可识别的行高／列宽对话框，当前网页界面需适配')

    def _close_dialog(self, dialog):
        cancel = self._unique_control(CANCEL_LABELS, dialog, optional=True)
        if cancel is not None:
            self._click(cancel, '取消')
        else:
            self.page.keyboard.press('Escape')
        for _ in range(20):
            if not dialog.is_visible():
                return
            self._pause()
        raise ValueError('表C尺寸对话框未能取消关闭，已停止')

    def _size_dialog(self, kind, native_unit=None):
        """打开行高/列宽设置对话框并解析单位。

        native_unit：对话框未标注单位时按Excel原生单位兜底
        （列宽=字符'excel_char'、行高=磅'pt'，WPS默认即此）；读回核验兜底。"""
        labels = COLUMN_LABELS if kind == 'column' else ROW_LABELS
        if self._visible_dialogs():
            raise ValueError('请先关闭表C现有弹窗，再重新上传')
        dialog = self._wps_inline_size(kind)
        if dialog is None:
            dialog = self._legacy_size_dialog(kind, labels)
        return self._parse_size_dialog(dialog, kind, labels, native_unit)

    def _wps_inline_size(self, kind):
        """Observed WPS canvas UI: whole-axis namebox selection + Shift+F10.

        Its role=option contains a numeric input, explicit unit and a tick,
        rather than a command that opens a separate OK/Cancel dialog. Selecting
        whole axes also preserves every column/row in a grouped dimension.
        No guessed canvas coordinates or private spreadsheet state are used.
        """
        canvases = [(frame, item) for frame in self._frames()
                    for item in frame.locator('canvas#et_canvas').all() if item.is_visible()]
        if not canvases:
            return None
        if len(canvases) != 1 or not self.last_ref:
            raise DimensionControlError('表C画布或尺寸选区无法唯一确认，已停止', {'kind':kind,'canvases':len(canvases)})
        first,last = self.last_ref.split(':')[0],self.last_ref.split(':')[-1]
        c1,r1 = address(first); c2,r2 = address(last)
        axis = f'{column_name(c1)}:{column_name(c2)}' if kind=='column' else f'{r1}:{r2}'
        label = '列宽' if kind=='column' else '行高'
        self.select_range(axis)
        self.progress(f'正在打开表C整{"列" if kind=="column" else "行"} {axis} 的{label}输入框…')
        self.page.keyboard.press('Shift+F10')
        found=[];match_count=0
        for _ in range(20):
            found=[];match_count=0
            for frame in self._frames():
                # Exact label in the observed option structure, never the
                # nearby insert/delete/autofit options or their number fields.
                # Do not retain an nth(index) locator: WPS inserts/reorders menu
                # options while loading. Resolve by its own visible label on
                # every subsequent read/click, including after React redraws.
                named=frame.locator('.et-dropdown-menu-span:visible').filter(
                    has_text=re.compile(r'^\s*'+re.escape(label)+r'\s*$'))
                options=frame.locator('[role="option"].ks-menu-item:visible').filter(has=named)
                count=options.count()
                match_count+=count
                if count:found.append(options)
            if match_count==1:
                option=found[0]
                ticks=option.locator('.et-input-menu-symbol-tick:visible')
                if ticks.count()==1:
                    return option
                break
            if match_count>1:break
            self._pause()
        inventory=self._control_inventory()
        screenshot=self._capture_ui_dump()
        self.page.keyboard.press('Escape')
        raise DimensionControlError(f'未找到表C的{label}直接输入框。请关闭登录成功提示，确认表C可编辑；当前批次尚未设置尺寸。',
            {'kind':kind,'axis_selection':axis,'matched_controls':match_count,
             'visible_controls':inventory,'screenshot':screenshot or ''})

    def _legacy_size_dialog(self, kind, labels):
        """Older non-canvas layouts retain their named-button/dialog path."""
        control = self._unique_control(labels, optional=True)
        if control is None:
            menu = self._unique_control(SIZE_MENUS, optional=True)
            if menu is not None:
                self._click(menu, SIZE_MENUS[0])
                control = self._unique_control(labels)
            else:
                # 部分WPS皮肤没有顶部“行和列”菜单：依次尝试快捷工具抽屉与右键表头。
                control = self._toolbar_size_item(kind, labels)
                if control is None:
                    control = self._context_size_item(kind, labels)
                if control is None:
                    # 右键可能直接弹出尺寸弹层：标准对话框或WPS自定义弹层都接受。
                    control = self._wait_popup() or None
                if control is None:
                    # 携带可见控件清单抛出，便于依据真实页面菜单适配；同时保存界面截图。
                    self._capture_ui_dump()
                    self._unique_control(labels)
        if control is not True:
            self._click(control, labels[0])
        return self._dialog()

    def _parse_size_dialog(self, dialog, kind, labels, native_unit):
        try:
            text = dialog.inner_text(timeout=3000)
            labelled = any(_label(name) in _label(text) for name in labels)
            if not labelled:
                # 右键表头直接弹出的对话框可能没有“列宽/行高”文字标题；
                # 必须同时具备确定与取消按钮才接受，避免吞掉无关弹窗。
                if (self._unique_control(CONFIRM_LABELS, dialog, optional=True) is None
                        or self._unique_control(CANCEL_LABELS, dialog, optional=True) is None):
                    evidence=dialog.evaluate('''el=>({role:el.getAttribute('role'),class:String(el.className||'').slice(0,160),
                        menu_labels:Array.from(el.querySelectorAll('.et-dropdown-menu-span')).slice(0,8).map(n=>(n.innerText||'').slice(0,60)),
                        inputs:el.querySelectorAll('input').length,ticks:el.querySelectorAll('.et-input-menu-symbol-tick').length})''')
                    raise DimensionControlError('表C当前弹窗不是所需的行高／列宽设置，已停止',
                        {'wanted':list(labels),'selected_range':self.last_ref,'detected_panel':evidence,
                         'screenshot':self._capture_ui_dump() or ''},code='DIMENSION_PANEL_MISMATCH')
            candidates = [item for item in dialog.locator('input:not([type="hidden"]),[role="spinbutton"][contenteditable="true"]').all()
                          if item.is_visible() and item.is_enabled() and item.is_editable()
                          and (item.get_attribute('type') or 'text') in ('text', 'number', 'search')]
            if len(candidates) != 1:
                raise ValueError('表C尺寸对话框的输入框无法唯一确认，已停止')
            field = candidates[0]
            detail = field.evaluate('''el => ({
                direct:[el.getAttribute('data-unit'),el.getAttribute('unit'),el.getAttribute('aria-label'),el.title,el.placeholder,
                    ...Array.from(el.labels||[]).map(label=>label.innerText)].filter(Boolean).join(' '),
                parent:el.parentElement ? el.parentElement.innerText : '',
                step:el.getAttribute('step'),value:el.value===undefined ? el.innerText : el.value
            })''')
            unit = None
            for evidence in (detail['direct'], detail['parent'], text):
                detected = _units(evidence)
                if len(detected) == 1:
                    unit = detected.pop()
                    break
                if len(detected) > 1:
                    raise ValueError('表C尺寸输入框同时出现多种单位，无法确定实际单位，已停止')
            if unit is None:
                if dialog.get_attribute('role') == 'option':
                    raise ValueError('表C菜单内尺寸输入框未明确标注单位，已停止')
                # 对话框未标注单位：按Excel原生单位处理（WPS默认），读回核验兜底。
                if native_unit in ('excel_char', 'pt'):
                    unit = native_unit
                else:
                    raise ValueError('表C尺寸输入框未明确标注像素、磅或字符单位，不能猜测单位；当前网页界面需适配')
            if kind == 'row' and unit == 'excel_char':
                raise ValueError('表C行高不能按字符单位设置，当前网页界面需适配')
            # A mixed selection may show an empty size.  Preflight may cancel
            # it; read-back verification below requires a concrete number.
            if str(detail['value']).strip():
                _numeric(detail['value'])
            step = None
            if detail['step'] not in (None, '', 'any'):
                step = _numeric(detail['step'])
                if step <= 0:
                    raise ValueError('表C尺寸输入框的步长无效')
            return dialog, field, unit, step
        except Exception:
            self._close_dialog(dialog)
            raise

    def _group_range(self, group, kind, target_range):
        parts = target_range.split(':')
        left, top = address(parts[0])
        right, bottom = address(parts[-1])
        start, end = group['start'], group['end']
        if kind == 'column':
            if not left <= start <= end <= right:
                raise ValueError('表C列宽设置范围超出上传范围')
            # 列宽作用于整列：只选目标范围的第一行，避免整表选区在网页超时。
            head = f'{column_name(start)}{top}'
            if start == end:
                return head  # WPS名称框不接受A1:A1这类单格范围
            return f'{head}:{column_name(end)}{top}'
        if not top <= start <= end <= bottom:
            raise ValueError('表C行高设置范围超出上传范围')
        # 行高作用于整行：只选目标范围的第一列，避免整表选区在网页超时。
        head = f'{column_name(left)}{start}'
        if start == end:
            return head
        return f'{head}:{column_name(left)}{end}'

    def _value_for_unit(self, group, kind, unit, step):
        value = _numeric(group['value'])
        if kind == 'row':
            target = value * 96 / 72 if unit == 'px' else value
        elif unit == 'excel_char':
            target = value
        else:
            if group.get('points') is None:
                raise ValueError('表C按像素／磅设置列宽，但缺少表C的Excel实测列宽，不能用字符宽度猜算')
            target = _numeric(group['points'])
            if unit == 'px':
                target *= 96 / 72
        if step is not None:
            target = math.floor(target / step + 0.5) * step
        return round(target, 6)

    def _sheet_tabs(self, name):
        matches, visible = [], []
        for frame in self._frames():
            result = frame.locator(SHEET_SELECTOR).evaluate_all(SHEET_SCAN, name)
            matches.extend((frame,item) for item in result['matches'])
            visible.extend({k:v for k,v in item.items() if k!='path'} for item in result['visible_tabs'])
        return matches, visible[:30]

    def select_sheet(self, name):
        if not isinstance(name, str) or not name.strip():
            raise ValueError('请填写表C工作表名称')
        wanted = name.strip()
        deadline = time.monotonic() + self.sheet_timeout
        while True:
            candidates, visible = self._sheet_tabs(wanted)
            if len(candidates)==1: break
            if time.monotonic() >= deadline:
                names = list(dict.fromkeys(label for item in visible for label in item['labels']))[:15]
                details = {'requested_sheet':wanted, 'visible_tabs':visible,
                           'matching_controls':len(candidates),
                           'frame_hosts':[urllib.parse.urlparse(f.url).hostname or 'local' for f in self.page.frames]}
                code = 'SHEET_TAB_AMBIGUOUS' if candidates else 'SHEET_TAB_NOT_FOUND'
                hint = '；当前可识别标签：'+ '、'.join(names) if names else '；页面尚未提供可识别的工作表标签'
                raise SheetTabError(code, f'无法唯一定位表C工作表标签“{wanted}”'+hint+'。请确认已登录、底部标签名称一致且表格加载完成；报告已记录标签识别结果', details)
            self.progress(f'正在等待表C工作表标签“{wanted}”加载（最多{self.sheet_timeout:g}秒）…')
            self.page.wait_for_timeout(400)
        frame, info = candidates[0]
        if not info['selected']:
            self._click(frame.locator(info['path']), '工作表标签')
        # SPA tab clicks can replace nodes. Query fresh DOM instead of retaining
        # a positional locator that could now refer to another worksheet.
        for _ in range(20):
            current, visible = self._sheet_tabs(wanted)
            if len(current)==1 and current[0][1]['selected']:
                self.progress(f'已确认表C当前工作表：{wanted}')
                return
            self._pause()
        raise SheetTabError('SHEET_TAB_UNCONFIRMED', f'表C工作表“{wanted}”未提供可确认的选中状态，已停止上传',
                            {'requested_sheet':wanted,'visible_tabs':visible})

    def preflight(self, plan):
        """Probe and cancel size dialogs before table C data is changed."""
        self.preflight_units = {}
        for kind, key in (('column', 'column_groups'), ('row', 'row_groups')):
            groups = plan.get(key) or []
            if not groups:
                continue
            self._select(self._group_range(groups[0], kind, plan['target_range']))
            dialog, field, unit, step = self._size_dialog(kind, 'excel_char' if kind == 'column' else 'pt')
            try:
                # Validate every group for this detected unit before upload,
                # including later columns that may lack measured point widths.
                for group in groups:
                    self._group_range(group, kind, plan['target_range'])
                    self._value_for_unit(group, kind, unit, step)
                self.preflight_units[kind] = unit
            finally:
                self._close_dialog(dialog)
        self.progress('已确认表C行高、列宽控件及单位，可以上传')

    def apply_dimensions(self, plan):
        total = sum(len(plan.get(key) or []) for key in ('column_groups', 'row_groups'))
        done = 0
        for kind, key in (('column', 'column_groups'), ('row', 'row_groups')):
            for group in plan.get(key) or []:
                ref = self._group_range(group, kind, plan['target_range'])
                self._select(ref)
                dialog, field, unit, step = self._size_dialog(kind, 'excel_char' if kind == 'column' else 'pt')
                try:
                    target = self._value_for_unit(group, kind, unit, step)
                    field.fill(format(target, '.6f').rstrip('0').rstrip('.') or '0', timeout=3000)
                    if dialog.get_attribute('role')=='option':
                        confirm=dialog.locator('.et-input-menu-symbol-tick:visible')
                        if confirm.count()!=1:raise ValueError('表C尺寸输入框的确认勾无法唯一定位')
                    else:
                        confirm = self._unique_control(CONFIRM_LABELS, dialog)
                    self._click(confirm, '确定')
                    for _ in range(20):
                        if not dialog.is_visible():
                            break
                        self._pause()
                    else:
                        raise ValueError('表C未接受行高／列宽设置，已停止；请检查页面提示')
                except Exception:
                    if dialog.is_visible():
                        self._close_dialog(dialog)
                    raise
                # Reopening the same control reads the resulting stored size.
                check, check_field, check_unit, _ = self._size_dialog(kind, unit)
                try:
                    if check_unit != unit:
                        raise ValueError('表C尺寸单位在设置后发生变化，无法核验')
                    raw = check_field.evaluate('el=>el.value===undefined ? el.innerText : el.value')
                    actual = _numeric(raw)
                    if not math.isclose(actual, target, rel_tol=0, abs_tol=0.01):
                        raise ValueError('表C行高／列宽读回结果与表C要求不一致，已停止')
                finally:
                    self._close_dialog(check)
                done += 1
                self.progress(f'正在规范表C行高／列宽：已核验{done}/{total}组')
        return dict(column_groups=len(plan.get('column_groups') or []),
                    row_groups=len(plan.get('row_groups') or []),verified_groups=done)

    def _merge_command(self, ref, unmerge):
        self._select(ref)
        labels = ('取消合并单元格', '取消合并', 'Unmerge Cells') if unmerge else ('合并所有', '合并单元格', 'Merge All', 'Merge Cells')
        # Merge is often a toggle button; only a menu item is accepted for that
        # ambiguous label.  Explicit Unmerge buttons have no toggle ambiguity.
        control = self._unique_control(labels, optional=True, role=None if unmerge else 'menuitem')
        if control is None:
            menu = self._unique_control(('合并选项', '合并单元格', '合并居中', '合并', 'Merge Options', 'Merge Cells'), menu_only=True)
            self._click(menu, '合并选项')
            control = self._unique_control(labels, role='menuitem')
        self._click(control, labels[0])
        if self._visible_dialogs():
            # Do not acknowledge any possible discard-data confirmation.
            self._close_dialog(self._dialog())
            raise ValueError('表C合并操作要求额外确认，已取消；请检查是否存在非左上角数据')

    def unmerge(self, ref):
        self._merge_command(ref, True)

    def merge(self, ref):
        self._merge_command(ref, False)
