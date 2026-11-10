"""Decode CF_HTML context and locate a complete table without rewriting it.

The operating-system clipboard is never changed. Only the validation view is
isolated from CF_HTML metadata and browser/Office document wrappers.
"""
from html.parser import HTMLParser
import re


class CopyIntegrityError(ValueError):
    def __init__(self, message, *, can_split=True, kind='html'):
        super().__init__(message)
        self.can_split=can_split
        # 'rows'：行数不足（缩小批次可能有用）；'columns'：列数/隐藏列推断不符；
        # 'html'：剪贴板本身不可用。调用方据此决定要不要重试、怎么重试。
        self.kind=kind


_HEADER=re.compile(rb'(Version|StartHTML|EndHTML|StartFragment|EndFragment|StartSelection|EndSelection|SourceURL):[^\r\n]*(?:\r\n|\r|\n|$)',re.I)


def decode_context(raw):
    """Offsets count UTF-8 bytes; whitespace, CR/LF, and -1 are valid headers."""
    if not isinstance(raw,(bytes,str)):
        raise CopyIntegrityError('剪贴板富文本类型无效（CF_HTML_TYPE），表A1未修改',can_split=False)
    if isinstance(raw,str) and not raw.lstrip('\ufeff').lower().startswith('version:'):
        return raw.lstrip('\ufeff').rstrip('\x00')
    data=raw.encode('utf-8') if isinstance(raw,str) else raw
    data=data.rstrip(b'\x00')
    position=3 if data.startswith(b'\xef\xbb\xbf') else 0
    offsets={}
    is_header=data[position:position+8].lower()==b'version:'
    if is_header:
        while True:
            match=_HEADER.match(data,position)
            if match is None:break
            key=match[1].decode('ascii').lower()
            if key not in ('version','sourceurl'):
                value=match[0].split(b':',1)[1].strip()
                if not re.fullmatch(rb'-?\d+',value) or key in offsets:
                    raise CopyIntegrityError('剪贴板偏移字段无效（CF_HTML_OFFSETS），表A1未修改',can_split=False)
                offsets[key]=int(value)
            position=match.end()
        pair=(offsets.get('starthtml'),offsets.get('endhtml'))
        if pair==(-1,-1):
            pair=(offsets.get('startfragment'),offsets.get('endfragment'))
        if None not in pair:
            first,last=pair
            if not position<=first<last<=len(data):
                raise CopyIntegrityError('剪贴板HTML字节范围无效或内容截断（CF_HTML_OFFSETS），表A1未修改',can_split=False)
            data=data[first:last]
        else:
            # Some applications omit offsets; consume only recognized header
            # lines instead of feeding them to the HTML parser as visible text.
            data=data[position:]
    try:
        return data.decode('utf-8-sig',errors='strict')
    except UnicodeDecodeError as exc:
        raise CopyIntegrityError('剪贴板HTML编码或字节偏移无效（CF_HTML_UTF8），表A1未修改',can_split=False) from exc


class _TableSlice(HTMLParser):
    """Locate real table tokens, ignoring lookalikes in comments/attributes/CSS."""
    def __init__(self, text):
        super().__init__(convert_charrefs=False)
        self.text=text
        self.start=self.stop=None
        self.active=False
        self.position_line=1
        self.position_offset=0

    def absolute_position(self):
        line,col=self.getpos()
        while self.position_line<line:
            self.position_offset=self.text.index('\n',self.position_offset)+1
            self.position_line+=1
        return self.position_offset+col

    def handle_starttag(self, tag, attrs):
        if tag in ('img','object','iframe'):
            raise CopyIntegrityError('复制内容含图片或嵌入对象，当前不能验证其完整性，表A1未修改',can_split=False)
        if tag=='table':
            if self.start is not None:
                raise CopyIntegrityError('剪贴板包含多个或嵌套表格，无法确认复制范围，表A1未修改',can_split=False)
            self.start=self.absolute_position()
            self.active=True
        elif tag in ('tr','td','th','thead','tbody','tfoot','col','colgroup') and not self.active:
            raise CopyIntegrityError('表格行列标签位于完整表格之外（HTML_TABLE_STRUCTURE），表A1未修改',can_split=False)

    def handle_endtag(self, tag):
        if tag=='table':
            if not self.active:
                raise CopyIntegrityError('表格结束标签没有对应起点（HTML_TABLE_STRUCTURE），表A1未修改',can_split=False)
            self.stop=self.text.index('>',self.absolute_position())+1
            self.active=False
        elif tag in ('tr','td','th','thead','tbody','tfoot','colgroup') and not self.active:
            raise CopyIntegrityError('表格行列结束标签位于表格之外（HTML_TABLE_STRUCTURE），表A1未修改',can_split=False)

    def handle_startendtag(self, tag, attrs):
        if tag in ('table','tr','td','th'):
            raise CopyIntegrityError('表格关键标签意外自闭合（HTML_TABLE_STRUCTURE），表A1未修改',can_split=False)
        self.handle_starttag(tag,attrs)

    def finish(self):
        super().close()
        if self.start is None:
            raise CopyIntegrityError('剪贴板未包含完整HTML表格（HTML_TABLE_MISSING），表A1未修改',can_split=False)
        if self.stop is None or self.active:
            raise CopyIntegrityError('富文本表格缺少结束标记，可能被截断；表A1未修改')
        return self.text[self.start:self.stop]


def table_fragment(raw):
    context=decode_context(raw)
    parser=_TableSlice(context)
    # Feed complete input once: unfinished large attributes should not be
    # rescanned for every tiny chunk. This parser builds no element tree.
    parser.feed(context)
    return parser.finish()
