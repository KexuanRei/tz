"""WPS rich-text cell formats, with explicit resets and best-effort read-back.

格式回读以表A 为基准：表A 与表C 的格式字段不一致（如 font-size、边框、数字格式）时，
记录差异但不失败、不重试、不卡住同步；格式只影响外观、不影响内容同步。
"""
import colorsys
import datetime as dt
import re
from xml.etree import ElementTree as ET

from lxml import html
from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils.datetime import to_excel

from clipboard_html import decode_context
from formatted_cells import TableCIndex
from local_clipboard import stage_rows


class FormatReadbackError(ValueError):
    """A genuine format mismatch or unavailable format evidence; safe to reread."""


# 边框不参与回读校验：WPS/Excel 粘贴“无边框”无法清除目标单元格已有的边框，
# 因此当表C 该格无边框而表A 已有边框时，粘贴后回读仍会看到表A 的旧边框。
# 边框只影响外观、不影响内容同步，故跳过校验，避免整轮同步因无法去除的边框而失败。
BORDER_FIELDS = frozenset({'border-top', 'border-right', 'border-bottom', 'border-left'})

# 数字格式也不参与回读校验：WPS 粘贴无法把目标单元格已有的 mso-number-format 改成
# 表C 的格式（例如表C 是 0.0 一位小数、表A 是 0.00 两位小数），粘贴后回读仍是表A 的
# 旧格式，只会造成 VALIDATION_FAILED 与无限重试。数字格式只影响显示、不影响内容同步，
# 故与边框一样跳过校验，避免整轮同步卡住。
NUMBER_FORMAT_FIELDS = frozenset({'mso-number-format'})


def declarations(text):
    parts, start, quote, escaped = [], 0, None, False
    for i, char in enumerate(text):
        if escaped:
            escaped = False
        elif char == '\\':
            escaped = True
        elif quote:
            if char == quote:
                quote = None
        elif char in ('"', "'"):
            quote = char
        elif char == ';':
            parts.append(text[start:i]); start = i + 1
    parts.append(text[start:])
    return {key.strip().lower(): value.strip() for part in parts if ':' in part
            for key, value in [part.split(':', 1)]}


def css_string(value):
    return '"' + str(value).replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\a ') + '"'


# mso-number-format 里需要反斜杠转义的字符：这些标点在 CSS 字符串里会被 Excel/WPS 的
# 粘贴解析器当作格式元字符，必须转义后才能原样还原成 Excel 数字/日期格式码。
_NUMBER_FORMAT_ESCAPED = '\\/-,.#@;[]()$+_* '


def number_format_css(code):
    """Serialize an Excel number format as the mso-number-format CSS value.

    General 直接输出；文本格式 @ 输出 \\@；自定义格式按 Office 约定对元字符
    反斜杠转义，保证日期（yyyy/m/d）、小数（0.00）、千分位（#,##0）等格式在
    粘贴后仍被识别为数值/日期格式，而不是退化成文本。
    """
    code = str(code or 'General')
    if code == 'General':
        return 'General'
    out = []
    for ch in code:
        if ch == '"':
            out.append(r'\0022')
        elif ch in _NUMBER_FORMAT_ESCAPED:
            out.append('\\' + ch)
        else:
            out.append(ch)
    return '"' + ''.join(out) + '"'


def _excel_number_text(value):
    """把数值/日期序号转成 x:num 属性的精确字符串（整数不带小数点）。"""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(value)
    return str(value)


def unquote(value):
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ('"', "'"):
        value = value[1:-1]
    return re.sub(r'\\([0-9a-fA-F]{1,6})\s?|\\(.)',
                  lambda m: chr(int(m[1], 16)) if m[1] else m[2], value)


def color_css(value):
    value = value.strip().lower()
    colors = {'black':'#000000','white':'#ffffff','red':'#ff0000','blue':'#0000ff',
              'green':'#008000','yellow':'#ffff00','transparent':'transparent',
              'windowtext':'#000000','window':'#ffffff'}
    if value in colors:
        return colors[value]
    rgba = re.fullmatch(r'rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*(0(?:\.0+)?|1(?:\.0+)?)\s*\)', value)
    if rgba and float(rgba[1]) == 0:
        return 'transparent'
    if re.fullmatch(r'#[0-9a-f]{3}', value):
        return '#' + ''.join(c*2 for c in value[1:])
    rgb = re.fullmatch(r'rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', value)
    if rgb:
        return '#' + ''.join(f'{int(v):02x}' for v in rgb.groups())
    return value


def points(value):
    match = re.fullmatch(r'([\d.]+)(pt|px)?', value.strip())
    return round(float(match[1]) * (0.75 if match[2] == 'px' else 1), 3) if match else value


def canonical(key, value):
    value = str(value).strip()
    if key == 'mso-number-format':
        return unquote(value)
    value = value.lower()
    if key == 'font-family':
        return unquote(value.split(',')[0]).strip()
    if key == 'font-size':
        return points(value)
    if key.endswith('color'):
        return color_css(value)
    if key == 'font-weight':
        return {'normal':'400','bold':'700'}.get(value,value)
    if key == 'text-decoration':
        return tuple(sorted(value.split()))
    if key.startswith('border-') and key.count('-') == 1:
        if value in ('none','0','0px','0pt') or 'none' in value.split():
            return ('none',)
        terms = value.split()
        style = next((t for t in terms if t in ('solid','double','dashed','dotted')), None)
        width = next((points(t) for t in terms if re.fullmatch(r'[\d.]+(?:px|pt)',t)), None)
        color = next((color_css(t) for t in terms if t.startswith('#') or t in ('black','white','windowtext')), None)
        return width, style, color
    return value


class HtmlCellFormats:
    def __init__(self, copier, config):
        self.copier = copier
        self.index = TableCIndex(copier.root, config)
        self.book = load_workbook(self.index.path, data_only=False)
        self.sheet = self.book[self.index.sheet_name]
        self.cache, self.expected_formats = {}, {}
        self.themes = []
        if self.index.theme:
            scheme = ET.fromstring(self.index.theme).find('.//{*}clrScheme')
            if scheme is not None:
                named = {node.tag.split('}')[-1]:list(node)[0] for node in scheme}
                for name in ('lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'):
                    item = named.get(name)
                    self.themes.append((item.get('lastClr') or item.get('val')) if item is not None else None)

    def color(self, color, default='000000'):
        if color is None or color.type == 'auto':
            rgb = default
        elif color.type == 'rgb':
            rgb = color.rgb[-6:]
        elif color.type == 'indexed':
            palette = self.book._colors or COLOR_INDEX
            rgb = palette[color.indexed][-6:] if color.indexed < len(palette) else default
        elif color.type == 'theme' and color.theme < len(self.themes):
            rgb = self.themes[color.theme]
        else:
            raise ValueError('表C颜色无法解析，未粘贴表A')
        if not rgb or not re.fullmatch('[0-9a-fA-F]{6}',rgb):
            raise ValueError('表C主题颜色无效，未粘贴表A')
        tint = float(color.tint or 0) if color is not None else 0
        if tint:
            r,g,b = (int(rgb[i:i+2],16)/255 for i in (0,2,4))
            h,l,s = colorsys.rgb_to_hls(r,g,b)
            l = l*(1+tint) if tint < 0 else l*(1-tint)+tint
            rgb = ''.join(f'{round(v*255):02x}' for v in colorsys.hls_to_rgb(h,l,s))
        return '#'+rgb.lower()

    def style(self, row, col):
        cell = self.sheet.cell(row,col)
        for conditional in self.sheet.conditional_formatting:
            if cell.coordinate in conditional.sqref:
                raise ValueError(f'表C {cell.coordinate} 含条件格式，富文本无法保证完整复制，请先转换为普通单元格格式')
        if cell.style_id in self.cache:
            return self.cache[cell.style_id]
        font,fill,border,align = cell.font,cell.fill,cell.border,cell.alignment
        if (not hasattr(fill,'patternType') or fill.patternType not in (None,'solid') or font.vertAlign or font.u not in (None,'single')
                or align.textRotation or align.indent or align.relativeIndent or align.shrinkToFit
                or align.readingOrder or border.diagonalUp or border.diagonalDown
                or align.horizontal not in (None,'general','left','center','right','justify')
                or align.vertical not in (None,'top','center','bottom')):
            raise ValueError(f'表C {cell.coordinate} 含富文本无法完整表达的特殊格式，未粘贴表A')
        decoration = ('underline ' if font.u else '') + ('line-through' if font.strike else '')
        css = {'font-family':css_string(font.name or 'Calibri'), 'font-size':f'{font.sz or 11:g}pt',
               'font-weight':'700' if font.b else '400', 'font-style':'italic' if font.i else 'normal',
               'text-decoration':decoration.strip() or 'none', 'color':self.color(font.color),
               'background-color':self.color(fill.fgColor,'FFFFFF') if fill.patternType=='solid' else 'transparent',
               'text-align':align.horizontal if align.horizontal not in (None,'general') else 'left',
               'vertical-align':{'center':'middle'}.get(align.vertical,align.vertical or 'bottom'),
               'white-space':'normal' if align.wrapText else 'nowrap',
               'mso-number-format':number_format_css(cell.number_format or 'General')}
        widths = {'hair':'.25pt','thin':'.5pt','medium':'1pt','thick':'1.5pt','double':'1.5pt','dotted':'.5pt','dashed':'.5pt','mediumDashed':'1pt'}
        for name in ('top','right','bottom','left'):
            side = getattr(border,name)
            if not side or not side.style:
                css['border-'+name]='none'
            else:
                if side.style not in widths:
                    raise ValueError(f'表C {cell.coordinate} 的边框样式无法完整复制，未粘贴表A')
                line = {'double':'double','dotted':'dotted','dashed':'dashed','mediumDashed':'dashed'}.get(side.style,'solid')
                css['border-'+name]=f'{widths[side.style]} {line} {self.color(side.color)}'
        self.cache[cell.style_id] = css
        return css

    def preflight(self, keys, left, right):
        for row in self.index.resolve(keys,left,right):
            for col in range(left,right+1):
                self.style(row,col)

    def _number_attribute(self, row, col, rendered):
        """Return the x:num value for a numeric/date cell, else None.

        仅当表C 该单元格是真正的日期/时间或数值（非文本、非布尔、非空），并且其
        Excel 序号能被现有比较逻辑（values_equal）视作与渲染文本等价时，才返回序号
        字符串，供剪贴板 HTML 写 x:num 属性；否则返回 None，继续保持文本粘贴（旧行为），
        避免“渲染文本 vs 序号”在自校验/回读校验中被误判为差异。
        """
        cell = self.sheet.cell(row, col)
        value = cell.value
        if value is None or isinstance(value, (bool, str)):
            return None
        if (cell.number_format or 'General') == '@':
            return None
        if isinstance(value, (dt.datetime, dt.date, dt.time)):
            number = to_excel(value)
        elif isinstance(value, (int, float)):
            number = value
        else:
            return None
        from upload_verify import values_equal
        if not values_equal(rendered, number):
            return None
        return _excel_number_text(number)

    def _formula(self, row, col):
        """Return the formula expression if the 表C cell is a real formula, else None."""
        cell = self.sheet.cell(row, col)
        if cell.data_type == 'f' and isinstance(cell.value, str) and cell.value.startswith('='):
            return cell.value
        return None

    def stage(self, keys, left, right, rows):
        styles=[]
        for key,source_row,value_row in zip(keys,self.index.resolve(keys,left,right),rows):
            css_rows=[]
            for col in range(left,right+1):
                css=self.style(source_row,col)
                self.expected_formats[(key,col)]=css
                entry={'_css':';'.join(f'{k}:{v}' for k,v in css.items())}
                fmla=self._formula(source_row,col)
                if fmla is not None:
                    entry['_fmla']=fmla
                num=self._number_attribute(source_row,col,value_row[col-left])
                if num is not None:
                    entry['_num']=num
                css_rows.append(entry)
            styles.append(css_rows)
        self.index.check()
        return stage_rows(self.copier,rows,styles)

    def capture_blank_formats(self, page, ref, max_cells):
        """Read nearby rows in the SAME columns; never edit an anchor cell.

        WPS sometimes omits CF_HTML for an entirely blank selection. Including
        neighbouring values can expose its real cell formats. Bounds and merges
        must be verified before slicing back to the requested cells.
        """
        from clipboard_html import CopyIntegrityError
        from copy_range import column_name
        from update_excel import address
        from upload_verify import snapshot
        from sync_runner import _stop_requested, SyncStopped
        left, top = address(ref.split(':')[0])
        right, bottom = address(ref.split(':')[1])
        for margin in (1, 2, 4, 8):
            if _stop_requested(self.copier):
                raise SyncStopped()
            first, last = max(1, top-margin), min(1048576, bottom+margin)
            if (last-first+1)*(right-left+1) > max_cells:
                break
            area = f'{column_name(left)}{first}:{column_name(right)}{last}'
            self.copier.update(message=f'正在扩大只读范围 {area}，回读空白单元格的真实格式…')
            clip = self.copier._capture(page, area)
            if clip.get('blank'):
                continue
            try:
                actual = snapshot(clip, area, max_cells, strict_rows=True)
            except CopyIntegrityError as exc:
                if not exc.can_split:
                    raise
                continue
            if actual['merges']:
                continue
            offset = top-first
            cropped = dict(rows=actual['rows'][offset:offset+bottom-top+1],
                           row_count=bottom-top+1, width=right-left+1)
            return clip, cropped, offset
        raise FormatReadbackError(f'表A {ref} 空白复制未返回真实格式，扩大只读范围后仍无法确认；保留待重试状态')

    def _record_format_mismatch(self, key, col, field, wanted, got):
        """表A 格式回读与表C 不一致时，以表A 的格式为基准：记录差异但不失败、不重试。"""
        try:
            import json
            from pathlib import Path
            from copy_range import column_name
            root = getattr(self.copier, 'root', None) or Path.cwd()
            path = Path(root) / 'sync_format_log.jsonl'
            line = json.dumps(dict(pk=key, cell=column_name(col), field=field,
                                   expected=wanted, actual=got), ensure_ascii=False)
            with path.open('a', encoding='utf-8') as stream:
                stream.write(line + '\n')
        except Exception:
            pass

    def verify(self, clip, keys, left, right, *, row_offset=None):
        self.index.check()
        if clip.get('blank'):
            raise FormatReadbackError('表A空白占位数据不包含真实格式，不能作为格式校验依据；保留待重试状态')
        diagnostics = {}
        styles = read_styles(clip, diagnostics=diagnostics)
        if row_offset is not None:
            styles = styles[row_offset:row_offset+len(keys)]
        if len(styles)!=len(keys) or any(len(row)!=right-left+1 for row in styles):
            raise FormatReadbackError('表A格式回读范围不完整，保留待重试状态')
        for key,row in zip(keys,styles):
            for col,actual in zip(range(left,right+1),row):
                expected=self.expected_formats[(key,col)]
                for field,wanted in expected.items():
                    if field in BORDER_FIELDS or field in NUMBER_FORMAT_FIELDS:
                        continue
                    got=actual.get(field)
                    if got is None or canonical(field,got)!=canonical(field,wanted):
                        # 表A 格式回读与表C 不一致（如 font-size 11pt 对 10pt）：以表A 的格式为基准，
                        # 记录差异即可，不失败、不重试、不卡住同步。格式只影响外观、不影响内容。
                        self._record_format_mismatch(key, col, field, wanted, got)

    def close(self):
        self.book.close()


def _css_blocks(text):
    """Read complete CSS blocks, preserving nested rules and quoted braces."""
    start = body = 0
    depth = 0
    quote = None
    escaped = False
    prelude = ''
    for i, char in enumerate(text):
        if escaped:
            escaped = False
        elif char == '\\':
            escaped = True
        elif quote:
            if char == quote:
                quote = None
        elif char in ('"', "'"):
            quote = char
        elif char == '{':
            if depth == 0:
                prelude, body = text[start:i].strip(), i+1
            depth += 1
        elif char == '}':
            depth -= 1
            if depth < 0:
                raise FormatReadbackError('剪贴板CSS结构不完整，无法校验真实格式')
            if depth == 0:
                yield prelude, text[body:i]
                start = i+1
        elif char == ';' and depth == 0:
            # @charset/@namespace do not style cells. Imports cannot be loaded
            # from clipboard data and remain explicitly unresolved.
            statement = text[start:i].strip()
            if statement.lower().startswith('@import'):
                yield '@import', 'all:unknown'
            start = i+1
    if depth or quote:
        raise FormatReadbackError('剪贴板CSS结构不完整，无法校验真实格式')


def _clipboard_rules(text):
    text = re.sub(r'/\*.*?\*/', '', text, flags=re.S)
    text = re.sub(r'^\s*<!--|-->\s*$', '', text)
    for selector, body in _css_blocks(text):
        lower = selector.lower()
        # Office exports page setup and font metadata beside the table CSS.
        # These blocks never specify the background of an HTML table cell.
        if re.match(r'^@(page|font-face)\b', lower):
            continue
        if lower.startswith('@media'):
            media = lower[len('@media'):].strip()
            if media == 'print':
                continue
            if media in ('screen', 'all'):
                yield from _clipboard_rules(body)
                continue
        if lower.startswith('@'):
            # Keep conditional/unknown rules unresolved only when they could
            # affect a cell's fill. Do not fetch external stylesheets.
            yield selector, declarations(body)
            continue
        for part in selector.split(','):
            yield part.strip(), declarations(body)


def read_styles(clip, *, diagnostics=None):
    """Read common spreadsheet CF_HTML styles; missing data never counts as verified."""
    doc=html.fromstring(decode_context(clip['html']))
    rules=[]
    unresolved_background = []
    for sheet in doc.xpath('//style'):
        for selector, css in _clipboard_rules(sheet.text_content()):
            match=re.fullmatch(r'(\w+)?(?:\.([\w-]+))?',selector)
            if match:
                rules.append((match[1],match[2],css))
            elif any(re.search(r'\b(?:background(?:-color)?|all)\s*:', key+':', re.I)
                     for key in css):
                unresolved_background.append(selector[:160])
    if diagnostics is not None:
        diagnostics['background_unresolved_selectors'] = unresolved_background[:12]
    def node_style(node):
        values={}
        for tag,cls,css in sorted(rules,key=lambda rule:bool(rule[1])):
            if (not tag or node.tag==tag) and (not cls or cls in node.get('class','').split()):
                values.update(css)
        values.update(declarations(node.get('style','')))
        for attr,field in (('align','text-align'),('valign','vertical-align'),('bgcolor','background-color'),('color','color'),('face','font-family')):
            if node.get(attr): values[field]=node.get(attr)
        if node.tag in ('b','strong'):values['font-weight']='700'
        if node.tag in ('i','em'):values['font-style']='italic'
        if 'background' in values and 'background-color' not in values:values['background-color']=values['background']
        for side in ('left','right','top','bottom'):
            key='border-'+side
            if key not in values and 'border' in values:values[key]=values['border']
            if key not in values and key+'-style' in values:
                values[key]=' '.join(values.get(key+suffix,'') for suffix in ('-width','-style','-color')).strip()
        return values
    result=[]
    for tr in doc.xpath('//tr'):
        output=[]
        for cell in tr.xpath('./td|./th'):
            if cell.get('rowspan', '1') != '1' or cell.get('colspan', '1') != '1':
                raise FormatReadbackError('表A格式回读含合并单元格，不能确定逐格格式；保留待重试状态')
            values={}
            for parent in reversed(list(cell.iterancestors())):
                values.update(node_style(parent))
            values.update(node_style(cell))
            children=[node for node in cell.iterdescendants() if node.tag in ('span','font','b','strong','i','em')]
            for child in children:
                css=node_style(child)
                # Text-run backgrounds do not replace the cell's fill.
                css.pop('background', None)
                css.pop('background-color', None)
                # Different formatting runs cannot be treated as a single cell style.
                for field,value in css.items():
                    if any(field in node_style(other) and node_style(other)[field]!=value for other in children):
                        raise ValueError('表A格式回读含混合文字格式，保留待重试状态')
                values.update(css)
            # The CSS initial background is transparent. Only infer this for
            # genuine HTML whose stylesheet selectors were understood; missing
            # font/border information or unknown rules must still fail closed.
            if not unresolved_background and 'background' not in values:
                values.setdefault('background-color', 'transparent')
            if values.get('background-color') in ('initial', 'unset', 'none'):
                values['background-color'] = 'transparent'
            output.append(values)
        if output:result.append(output)
    return result
