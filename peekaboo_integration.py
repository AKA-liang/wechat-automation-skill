# peekaboo_integration.py
# 封装 PeekabooWin CLI 调用，解析 snapshot JSON

import subprocess
import json
import os
import re
import time
import ctypes
from pathlib import Path
from typing import Optional, List, Dict, Any
import pyautogui
pyautogui.FAILSAFE = False
import pyperclip


# PeekabooWin 路径检测顺序
PEEKABOO_DEV_PATH = r"C:\Users\13265\Desktop\here\PeekabooWin-main\PeekabooWin-main"
PEEKABOO_BIN = "bin\\peekaboo-win.js"


def get_peekaboo_path() -> Optional[str]:
    """检测 PeekabooWin 安装路径"""
    dev_bin = os.path.join(PEEKABOO_DEV_PATH, PEEKABOO_BIN)
    if os.path.exists(dev_bin):
        return PEEKABOO_DEV_PATH
    try:
        result = subprocess.run(
            ["npm", "root", "-g"], capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            global_path = os.path.join(result.stdout.strip(), "peekaboo-win", PEEKABOO_BIN)
            if os.path.exists(global_path):
                return os.path.join(result.stdout.strip(), "peekaboo-win")
    except Exception:
        pass
    return None


PEEKABOO_PATH = get_peekaboo_path()
PEEKABOO_EXE = os.path.join(PEEKABOO_PATH, PEEKABOO_BIN) if PEEKABOO_PATH else None


def peekaboo(args: str, timeout: int = 30) -> str:
    """执行 peekaboo-win 命令，返回 stdout"""
    if not PEEKABOO_EXE:
        raise RuntimeError("PeekabooWin not found")
    cmd = f'node "{PEEKABOO_EXE}" {args}'
    result = subprocess.run(
        cmd, shell=True, capture_output=True, text=True, timeout=timeout, encoding="utf-8", errors="replace"
    )
    return result.stdout + result.stderr


def peekaboo_see(title: str = "微信", mode: str = "window") -> Dict[str, Any]:
    """执行 peekaboo see，返回 parsed snapshot 数据"""
    output = peekaboo(f'see --mode {mode} --title "{title}"')
    match = re.search(r'"snapshotId":\s*"([^"]+)"', output)
    if not match:
        raise RuntimeError(f"Cannot get snapshotId from peekaboo output:\n{output}")
    snapshot_id = match.group(1)
    snapshot_file = Path(os.path.expanduser("~")) / ".peekaboo-windows" / "snapshots" / snapshot_id / "snapshot.json"
    if not snapshot_file.exists():
        raise RuntimeError(f"Snapshot file not found: {snapshot_file}")
    with open(snapshot_file, "r", encoding="utf-8", errors="replace") as f:
        data = json.load(f)
    return {
        "snapshotId": snapshot_id,
        "rawOutput": output,
        "data": data["result"],
    }


def click_on_text(text: str, snapshot_id: str = "latest") -> tuple:
    """用 OCR 文本匹配方式点击元素"""
    output = peekaboo(f'click --on "{text}" --snapshot {snapshot_id}')
    ok = '"element":' in output or '"ok": true' in output.lower() or 'No snapshot label target' not in output
    return ok, output[:300]


def click_by_element_id(snapshot_id: str, element_id: str) -> str:
    """通过 element ID 点击"""
    return peekaboo(f'snapshot click --snapshot {snapshot_id} --element-id {element_id}')


def click_coordinates(x: int, y: int) -> str:
    """点击绝对坐标"""
    return peekaboo(f'mouse click --x {x} --y {y}')


def type_text(text: str) -> str:
    """输入文字"""
    return peekaboo(f'type --text "{text}"')


def press_keys(keys: str) -> str:
    """按键"""
    return peekaboo(f'press --keys "{keys}"')


def focus_window(title: str) -> str:
    """聚焦窗口"""
    return peekaboo(f'window focus --title "{title}"')


# --- 新的搜索流程函数 ---

def get_wechat_hwnd():
    """获取微信窗口句柄"""
    return ctypes.windll.user32.FindWindowW("Qt51514QWindowIcon", "微信")


def maximize_wechat():
    """最大化微信窗口"""
    hwnd = get_wechat_hwnd()
    if hwnd:
        ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE = 3
        time.sleep(0.3)


def search_and_open_contact(contact_name: str, timeout: int = 30) -> bool:
    """
    使用 Ctrl+F 搜索流程打开联系人聊天窗口。
    流程：最大化窗口 → Ctrl+F 激活搜索 → 粘贴联系人名 → Enter 打开聊天
    
    Args:
        contact_name: 联系人名称（会被粘贴到搜索框）
        timeout: 超时时间（秒）
    
    Returns:
        bool: 是否成功打开聊天窗口
    """
    hwnd = get_wechat_hwnd()
    if not hwnd:
        return False

    # Step 1: 最大化窗口（让坐标固定）
    ctypes.windll.user32.ShowWindow(hwnd, 3)
    time.sleep(0.3)

    # Step 2: 聚焦窗口并激活搜索
    ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(0.3)
    press_keys("^f")
    time.sleep(0.5)

    # Step 3: 粘贴联系人名字（剪贴板比 peekaboo type 更可靠）
    pyperclip.copy(contact_name)
    time.sleep(0.1)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.8)

    # Step 4: 按回车打开第一个搜索结果
    press_keys("{enter}")
    time.sleep(2)

    # Step 5: 验证聊天窗口已打开（OCR 中有 >= 2 个联系人名）
    try:
        snapshot = peekaboo_see(title="微信", mode="window")
        lines = snapshot["data"].get("ocr", {}).get("lines", [])
        contact_lower = contact_name.lower()
        count = sum(1 for l in lines if contact_lower in l.get("text", "").lower())
        return count >= 2
    except Exception:
        return False


def send_message_via_clipboard(message: str) -> bool:
    """
    通过剪贴板粘贴 + 回车发送消息。
    调用前需确保聊天窗口已打开并聚焦。
    """
    try:
        pyperclip.copy(message)
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.3)
        press_keys("{enter}")
        time.sleep(1)
        return True
    except Exception:
        return False