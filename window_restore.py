# window_restore.py
# Win32 API 窗口恢复 - 处理微信窗口被最小化/隐藏到屏幕外的问题

import ctypes
from ctypes import wintypes

SW_RESTORE = 9


class RECT(ctypes.Structure):
    _fields_ = [
        ("Left", wintypes.LONG),
        ("Top", wintypes.LONG),
        ("Right", wintypes.LONG),
        ("Bottom", wintypes.LONG),
    ]


class Win32Restore:
    """Win32 窗口恢复工具"""

    _kernel32 = ctypes.windll.kernel32
    _user32 = ctypes.windll.user32

    FindWindow = _user32.FindWindowW
    ShowWindow = _user32.ShowWindow
    SetForegroundWindow = _user32.SetForegroundWindow
    GetWindowRect = _user32.GetWindowRect
    MoveWindow = _user32.MoveWindow

    @classmethod
    def find_wechat_window(cls):
        """找到微信窗口句柄"""
        # 方法1: FindWindow by class + title
        hwnd = cls.FindWindow("Qt51514QWindowIcon", "微信")
        if hwnd:
            return hwnd

        # 方法2: 枚举窗口查找
        results = []

        def enum_callback(hwnd, _):
            if cls._user32.IsWindowVisible(hwnd):
                length = cls._user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buf = ctypes.create_unicode_buffer(length + 1)
                    cls._user32.GetWindowTextW(hwnd, buf, length + 1)
                    title = buf.value
                    if "微信" in title:
                        results.append(hwnd)
            return True

        cls._user32.EnumWindows(ctypes.WINFUNCTYPE(
            ctypes.c_bool, wintypes.HWND, ctypes.c_void_p
        )(enum_callback), 0)

        return results[0] if results else None

    @classmethod
    def get_window_rect(cls, hwnd):
        """获取窗口位置和大小"""
        rect = RECT()
        cls.GetWindowRect(hwnd, ctypes.byref(rect))
        return {
            "left": rect.Left,
            "top": rect.Top,
            "right": rect.Right,
            "bottom": rect.Bottom,
            "width": rect.Right - rect.Left,
            "height": rect.Bottom - rect.Top,
        }

    @classmethod
    def is_window_off_screen(cls, rect):
        """判断窗口是否在屏幕外或大小异常"""
        return (
            rect["left"] < -10000 or
            rect["top"] < -10000 or
            rect["width"] < 200 or
            rect["height"] < 200
        )

    @classmethod
    def restore_window(cls, hwnd):
        """恢复微信窗口到前台可见位置"""
        rect = cls.get_window_rect(hwnd)

        if cls.is_window_off_screen(rect):
            cls.ShowWindow(hwnd, SW_RESTORE)
            ctypes.windll.kernel32.Sleep(300)
            cls.MoveWindow(hwnd, 100, 100, 1118, 809, True)
            ctypes.windll.kernel32.Sleep(500)
        else:
            cls.ShowWindow(hwnd, SW_RESTORE)
            ctypes.windll.kernel32.Sleep(300)

        cls.SetForegroundWindow(hwnd)
        ctypes.windll.kernel32.Sleep(300)

        return cls.get_window_rect(hwnd)


def restore_wechat_window():
    """恢复微信窗口，返回 (成功标志, 窗口信息)"""
    hwnd = Win32Restore.find_wechat_window()
    if not hwnd:
        return False, "微信窗口未找到"

    rect = Win32Restore.restore_window(hwnd)
    return True, {
        "hwnd": hwnd,
        "position": {"x": rect["left"], "y": rect["top"]},
        "size": {"width": rect["width"], "height": rect["height"]},
    }