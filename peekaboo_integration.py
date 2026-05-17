# peekaboo_integration.py
# 封装 PeekabooWin CLI 调用，解析 snapshot JSON
# 继承方案B的 UI 树定位能力，但用 Python 实现

import subprocess
import json
import os
import re
from pathlib import Path
from typing import Optional, Dict, Any


# PeekabooWin 路径检测顺序
PEEKABOO_DEV_PATH = r"C:\Users\13265\Desktop\here\PeekabooWin-main\PeekabooWin-main"
PEEKABOO_BIN = "bin\\peekaboo-win.js"


def get_peekaboo_path() -> Optional[str]:
    """检测 PeekabooWin 安装路径"""
    # 优先开发路径
    dev_bin = os.path.join(PEEKABOO_DEV_PATH, PEEKABOO_BIN)
    if os.path.exists(dev_bin):
        return PEEKABOO_DEV_PATH

    # 全局 npm
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
        raise RuntimeError("PeekabooWin not found. Install via: npm install -g peekaboo-win")

    cmd = f'node "{PEEKABOO_EXE}" {args}'
    result = subprocess.run(
        cmd, shell=True, capture_output=True, text=True, timeout=timeout, encoding="utf-8", errors="replace"
    )
    return result.stdout + result.stderr


def peekaboo_see(title: str = "微信", mode: str = "window") -> Dict[str, Any]:
    """执行 peekaboo see，返回 parsed snapshot 数据"""
    output = peekaboo(f'see --mode {mode} --title "{title}"')

    # 提取 snapshotId
    match = re.search(r'"snapshotId":\s*"([^"]+)"', output)
    if not match:
        raise RuntimeError(f"Cannot get snapshotId from peekaboo output:\n{output}")

    snapshot_id = match.group(1)

    # 读取 snapshot JSON 文件（结果在 result 字段下）
    snapshot_file = Path(os.path.expanduser("~")) / ".peekaboo-windows" / "snapshots" / snapshot_id / "snapshot.json"
    if not snapshot_file.exists():
        raise RuntimeError(f"Snapshot file not found: {snapshot_file}")

    with open(snapshot_file, "r", encoding="utf-8", errors="replace") as f:
        data = json.load(f)

    return {
        "snapshotId": snapshot_id,
        "rawOutput": output,
        "data": data["result"],  # 实际数据在 result 字段下
    }


def click_on_text(text: str, snapshot_id: str = "latest") -> tuple:
    """用 OCR 文本匹配方式点击元素（contact 名通过 OCR 识别）"""
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