"""Retry helpers for transient Excel/COM and clipboard errors.

Excel raises "Call was rejected by callee" (RPC_E_CALL_REJECTED) while it is busy,
and the clipboard can be locked for a few milliseconds by another application.
Both are temporary and must not abort a long upload.
"""
import time

BUSY_HRESULTS = {
    0x80010001,   # RPC_E_CALL_REJECTED  (-2147418111)
    0x8001010A,   # RPC_E_SERVERCALL_RETRYLATER (-2147417846)
    0x80010100,   # RPC_E_SYS_CALL_FAILED
    0x80040003,   # clipboard locked while opening
    0x800401D0,   # CLIPBRD_E_CANT_OPEN
}
BUSY_TEXT = ('call was rejected', 'server busy', 'rpc', 'busy', 'openclipboard',
             'clipboard', '拒绝', '忙')


def _transient(exc):
    code = getattr(exc, 'hresult', None)
    if code is None:
        code = getattr(exc, 'winerror', None)
    if isinstance(code, int) and (code & 0xFFFFFFFF) in BUSY_HRESULTS:
        return True
    text = str(exc).lower()
    return any(word in text for word in BUSY_TEXT)


def retry_call(function, attempts=6, delay=0.35):
    """Call function(), retrying only while the error looks temporary."""
    for attempt in range(attempts):
        try:
            return function()
        except Exception as exc:
            if attempt == attempts - 1 or not _transient(exc):
                raise
            time.sleep(delay * (attempt + 1))
