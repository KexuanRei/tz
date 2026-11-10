"""Put a local mirror rectangle on the Windows clipboard as literal values."""
import csv
import html
import io


def _html_cell(value, style=None):
    # Keep the Office string value and visible text identical. Some spreadsheet
    # importers use x:str as the value, so x:str="" can erase a nonempty cell.
    escaped = html.escape(value, quote=True)
    attribute = escaped.replace('\r', '&#13;').replace('\n', '&#10;').replace('\t', '&#9;')
    display = escaped.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '<br>')
    css = style_to_css(style) if style else "mso-number-format:'\\@'"
    # 公式单元格用 x:fmla 携带公式表达式本身（显示文本仍写公式），粘贴后是活公式，
    # 不再退化为纯文本/纯数值。
    fmla = (style or {}).get('_fmla') if isinstance(style, dict) else None
    if fmla is not None:
        return ('<td x:str="' + attribute + '" x:fmla="' + html.escape(str(fmla), quote=True)
                + '" style="' + html.escape(css, quote=False).replace('"', '&quot;')
                + '">' + display + '</td>')
    # 数字/日期单元格用 x:num 携带精确数值（日期为 Excel 序号），并配 mso-number-format
    # 保留数值格式与日期格式；纯文本仍用 x:str，避免被 Excel 重新解析成数值/日期。
    num = (style or {}).get('_num') if isinstance(style, dict) else None
    if num is not None:
        return ('<td x:num="' + html.escape(str(num), quote=True) + '" style="'
                + html.escape(css, quote=False).replace('"', '&quot;') + '">' + display + '</td>')
    return '<td x:str="' + attribute + '" style="' + html.escape(css, quote=False).replace('"', '&quot;') + '">' + display + '</td>'


_BORDER_WIDTH = {
    'thin': '.5pt', 'hair': '.25pt', 'dashed': '.5pt', 'dotted': '.5pt',
    'dashDot': '.5pt', 'dashDotDot': '.5pt', 'slantDashDot': '.5pt',
    'medium': '1pt', 'mediumDashed': '1pt', 'mediumDashDot': '1pt',
    'mediumDashDotDot': '1pt', 'thick': '1.5pt', 'double': '1.5pt',
}
_BORDER_STYLE = {
    'dashed': 'dashed', 'dotted': 'dotted', 'double': 'double',
    'dashDot': 'dashed', 'dashDotDot': 'dashed', 'mediumDashed': 'dashed',
    'mediumDashDot': 'dashed', 'mediumDashDotDot': 'dashed', 'slantDashDot': 'dashed',
}
_VERTICAL = {'top': 'top', 'center': 'middle', 'bottom': 'bottom'}


def _number(value):
    return ('%g' % float(value)) + 'pt'


def _number_format_css(code):
    if code in (None, '', 'General'):
        return 'mso-number-format:General'
    if code == '@':
        return r"mso-number-format:'\@'"
    escaped = str(code).replace('\\', '\\\\').replace("'", "\\'")
    return "mso-number-format:'" + escaped + "'"


def style_to_css(style):
    """Serialize one flat style dict (see table_c_styles.extract_style) to inline CSS.

    Value stays a literal string (the ``x:str`` attribute); the style only
    changes how the cell is displayed. Properties that are absent keep their
    defaults in the pasted cell.
    """
    if not style:
        return ''
    if '_css' in style:
        return style['_css']
    parts = []
    if style.get('font_name'):
        parts.append('font-family:' + style['font_name'])
    if style.get('font_size'):
        parts.append('font-size:' + _number(style['font_size']))
    if style.get('font_bold') is not None:
        parts.append('font-weight:' + ('700' if style['font_bold'] else '400'))
    if style.get('font_italic') is not None:
        parts.append('font-style:' + ('italic' if style['font_italic'] else 'normal'))
    decorations = []
    if style.get('font_underline'):
        decorations.append('underline')
    if style.get('font_strike'):
        decorations.append('line-through')
    if decorations:
        parts.append('text-decoration:' + ' '.join(decorations))
    if style.get('font_color'):
        parts.append('color:#' + style['font_color'])
    if style.get('fill_color'):
        parts.append('background-color:#' + style['fill_color'])
    for side in ('top', 'right', 'bottom', 'left'):
        border = style.get('border_' + side)
        if not border:
            continue
        css_style = _BORDER_STYLE.get(border, 'solid')
        width = _BORDER_WIDTH.get(border, '.5pt')
        css = 'border-%s:%s %s' % (side, width, css_style)
        color = style.get('border_' + side + '_color')
        if color:
            css += ' #' + color
        parts.append(css)
    if style.get('align_horizontal'):
        parts.append('text-align:' + style['align_horizontal'])
    if style.get('align_vertical') in _VERTICAL:
        parts.append('vertical-align:' + _VERTICAL[style['align_vertical']])
    if style.get('wrap_text') is not None:
        parts.append('white-space:' + ('normal' if style['wrap_text'] else 'nowrap'))
    parts.append(_number_format_css(style.get('number_format')))
    return ';'.join(parts)


def _normalize_styles(rows, styles):
    if styles is None:
        return [[None] * len(row) for row in rows]
    if len(styles) != len(rows):
        raise ValueError('表A1复制格式与内容行数不一致')
    normalized = []
    for row, style_row in zip(rows, styles):
        if style_row is None:
            normalized.append([None] * len(row))
        elif len(style_row) != len(row):
            raise ValueError('表A1复制格式与内容列数不一致')
        else:
            normalized.append(list(style_row))
    return normalized


def clipboard_rows(rows, styles=None):
    if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
        raise ValueError('表A1复制范围必须为非空矩形')
    rows = [['' if v is None else str(v) for v in row] for row in rows]
    if any('\x00' in v for row in rows for v in row):
        raise ValueError('表A1内容含空字符，无法完整写入剪贴板，已停止粘贴')
    style_rows = _normalize_styles(rows, styles)
    text = io.StringIO(newline='')
    csv.writer(text, delimiter='\t', lineterminator='\r\n').writerows(rows)
    fragment = '<table>' + ''.join('<tr>' + ''.join(
        _html_cell(v, s)
        for v, s in zip(row, style_row)) + '</tr>' for row, style_row in zip(rows, style_rows)) + '</table>'
    prefix = '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><!--StartFragment-->'
    suffix = '<!--EndFragment--></body></html>'
    header = 'Version:0.9\r\nStartHTML:{:010d}\r\nEndHTML:{:010d}\r\nStartFragment:{:010d}\r\nEndFragment:{:010d}\r\n'
    start = len(header.format(0, 0, 0, 0).encode('ascii'))
    a = start + len(prefix.encode('utf-8'))
    b = a + len(fragment.encode('utf-8'))
    payload = (header.format(start, b+len(suffix.encode('utf-8')), a, b) + prefix+fragment+suffix).encode('utf-8')
    return dict(text=text.getvalue(), html=payload, styles=style_rows)


def stage_rows(copier, rows, styles=None):
    clip = clipboard_rows(rows, styles)
    # Browser/clipboard operations remain on the existing worker thread.
    copier._set_local_clipboard(rows, clip)
    return clip


def verify_clipboard(clip):
    """Check the OS clipboard, not just the in-memory payload used to write it."""
    from clipboard_html import decode_context
    from excel_clipboard import read_clipboard
    actual = read_clipboard()
    if (actual['text'] != clip['text'] or
            decode_context(actual['html']) != decode_context(clip['html']) or
            ('sequence' in clip and actual['sequence'] != clip['sequence'])):
        raise ValueError('粘贴前剪贴板内容缺失或已被改写，已停止，未执行本次表A粘贴')
    return actual['sequence']


def write_clipboard(clip):
    import win32clipboard as cb
    import win32con
    # pywin32 312 on Python 3.10 can corrupt the native heap when
    # SetClipboardData converts a Python str. Supply CF_UNICODETEXT as an
    # explicitly terminated UTF-16 buffer instead. memoryview also bypasses
    # pywin32's bytes-specific, one-byte text terminator handling.
    text_buffer = memoryview(clip['text'].encode('utf-16-le') + b'\x00\x00')
    cb.OpenClipboard()
    try:
        cb.EmptyClipboard()
        cb.SetClipboardData(win32con.CF_UNICODETEXT, text_buffer)
        cb.SetClipboardData(cb.RegisterClipboardFormat('HTML Format'), clip['html'])
    finally:
        cb.CloseClipboard()
