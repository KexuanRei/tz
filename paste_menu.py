"""Dismiss WPS paste options without applying another paste mode."""
from online_format import OnlineFormatter


def dismiss_paste_options(page):
    """Only Escape a visible paste menu, not arbitrary dialogs or cell text."""
    def visible():
        for frame in OnlineFormatter(page, lambda _: None)._frames():
            for item in frame.get_by_text('粘贴值和数字格式', exact=True).all():
                if item.is_visible() and item.evaluate('''el => {
                    if (el.closest('[role="grid"],[role="gridcell"],td,th')) return false;
                    for (let node = el; node && node !== document.body; node = node.parentElement) {
                        const text = node.innerText || '';
                        if (node.matches('[role="menu"],[role="listbox"],[class*="menu" i],[class*="popup" i],[class*="dropdown" i]') &&
                            text.includes('粘贴值') && text.includes('粘贴格式')) return true;
                    }
                    return false;
                }'''):
                    return True
        return False
    for _ in range(2):
        if not visible():
            return
        page.keyboard.press('Escape')
        page.wait_for_timeout(100)
    if visible():
        raise ValueError('表A粘贴选项菜单未关闭，已停止后续操作；请关闭菜单后重试')
