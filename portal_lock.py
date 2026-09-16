"""
Cross-process lock for Bloomberg Portal browser automation.

portal_register_serial.py / portal_pack_post.py / portal_download_labels.py /
portal_update_pod.py all connect to the SAME shared Chrome tab via
`playwright.chromium.connect_over_cdp(...)` and `context.pages[-1]`. If two of
these scripts run at the same time (e.g. two "1. Serial 등록 & QR 인쇄" clicks
back to back before the first finishes), they grab the same tab object and can
navigate/click on top of each other mid-action - this has actually happened
live (2026-05-28: a second order's Serial registration ran against the first
order's still-open delivery page and downloaded/printed the wrong label).

Wrap the full CDP session of each script in `portal_browser_lock()` so only
one such script ever touches the shared tab at a time; the rest wait their
turn instead of racing.
"""

import contextlib
import os
import time

BASE_DIR = os.path.dirname(__file__)
LOCK_PATH = os.path.join(BASE_DIR, "portal_browser.lock")

STALE_SECONDS = 300   # no single portal step should ever legitimately take this long
WAIT_TIMEOUT = 180     # how long to wait for the other script to finish before giving up
POLL_INTERVAL = 1


class PortalBusyError(RuntimeError):
    pass


def _clear_if_stale():
    try:
        age = time.time() - os.path.getmtime(LOCK_PATH)
    except FileNotFoundError:
        return
    if age > STALE_SECONDS:
        try:
            os.remove(LOCK_PATH)
        except FileNotFoundError:
            pass


@contextlib.contextmanager
def portal_browser_lock(label=""):
    """Block until the shared Chrome tab is free, then hold it exclusively."""
    _clear_if_stale()
    deadline = time.time() + WAIT_TIMEOUT
    fd = None
    while fd is None:
        try:
            fd = os.open(LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            _clear_if_stale()
            if time.time() > deadline:
                raise PortalBusyError(
                    "다른 Portal 자동화 작업이 아직 실행 중입니다 (같은 Chrome 탭을 공유합니다). "
                    "그 작업이 끝난 뒤 다시 시도하세요."
                )
            time.sleep(POLL_INTERVAL)
    try:
        os.write(fd, f"pid={os.getpid()} label={label}".encode("utf-8"))
        os.close(fd)
        yield
    finally:
        try:
            os.remove(LOCK_PATH)
        except FileNotFoundError:
            pass
