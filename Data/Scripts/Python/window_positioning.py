import ctypes
from ctypes import wintypes


MONITOR_DEFAULTTONEAREST = 2


class POINT(ctypes.Structure):
    _fields_ = [
        ("x", wintypes.LONG),
        ("y", wintypes.LONG),
    ]


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", wintypes.LONG),
        ("top", wintypes.LONG),
        ("right", wintypes.LONG),
        ("bottom", wintypes.LONG),
    ]


class MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("rcMonitor", RECT),
        ("rcWork", RECT),
        ("dwFlags", wintypes.DWORD),
    ]


def _get_monitor_work_area_from_point(x, y):
    user32 = ctypes.windll.user32

    pt = POINT(x, y)
    monitor = user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
    if not monitor:
        return None

    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(MONITORINFO)

    success = user32.GetMonitorInfoW(monitor, ctypes.byref(info))
    if not success:
        return None

    return (
        info.rcWork.left,
        info.rcWork.top,
        info.rcWork.right,
        info.rcWork.bottom,
    )


def _get_cursor_position():
    pt = POINT()
    if ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)):
        return pt.x, pt.y
    return 0, 0


def center_window(window):
    window.update_idletasks()

    win_w = window.winfo_width()
    win_h = window.winfo_height()

    if win_w <= 1:
        win_w = window.winfo_reqwidth()
    if win_h <= 1:
        win_h = window.winfo_reqheight()

    cursor_x, cursor_y = _get_cursor_position()
    work_area = _get_monitor_work_area_from_point(cursor_x, cursor_y)

    if work_area is None:
        screen_w = window.winfo_screenwidth()
        screen_h = window.winfo_screenheight()
        left, top, right, bottom = 0, 0, screen_w, screen_h
    else:
        left, top, right, bottom = work_area

    area_w = right - left
    area_h = bottom - top

    x = left + ((area_w - win_w) // 2)
    y = top + ((area_h - win_h) // 2)

    if x < left:
        x = left
    if y < top:
        y = top

    window.geometry(f"{win_w}x{win_h}+{x}+{y}")


def center_child_window(window, parent):
    window.update_idletasks()
    parent.update_idletasks()

    win_w = window.winfo_width()
    win_h = window.winfo_height()

    if win_w <= 1:
        win_w = window.winfo_reqwidth()
    if win_h <= 1:
        win_h = window.winfo_reqheight()

    parent_x = parent.winfo_rootx()
    parent_y = parent.winfo_rooty()
    parent_w = parent.winfo_width()
    parent_h = parent.winfo_height()

    x = parent_x + ((parent_w - win_w) // 2)
    y = parent_y + ((parent_h - win_h) // 2)

    window.geometry(f"{win_w}x{win_h}+{x}+{y}")