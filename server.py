# server.py
# 整合后的 WeChat MCP Server
# 核心思路：最大化窗口 + Ctrl+F 搜索流程打开联系人，剪贴板粘贴发送消息

import asyncio
import json
import sys
import time

import pyautogui
pyautogui.FAILSAFE = False
import pyperclip

from window_restore import restore_wechat_window
from peekaboo_integration import (
    peekaboo,
    peekaboo_see,
    click_on_text,
    click_by_element_id,
    click_coordinates,
    type_text,
    press_keys,
    focus_window,
    PEEKABOO_PATH,
    search_and_open_contact,
    send_message_via_clipboard,
    maximize_wechat,
    get_wechat_hwnd,
)


# ========== 核心发送流程 ==========

def send_message_to_contact(contact_name: str, message: str) -> tuple:
    """给指定联系人发送消息（完整流程）"""
    print(f"[Send] Target: {contact_name}, Message: {message}")

    # Step 1: 搜索并打开联系人聊天窗口
    print(f"\n[1/4] Searching and opening chat with {contact_name}...")
    ok = search_and_open_contact(contact_name)
    if not ok:
        return False, f"无法打开与 {contact_name} 的聊天窗口"

    # Step 2: 聚焦窗口
    print("\n[2/4] Focusing window...")
    hwnd = get_wechat_hwnd()
    if hwnd:
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(0.3)

    # Step 3: 点击输入框（用 OCR 找"输入"文字）
    print("\n[3/4] Clicking input field...")
    ok, _ = click_on_text("输入", snapshot_id="latest")
    if not ok:
        # 降级：窗口中央下方坐标
        try:
            snapshot = peekaboo_see(title="微信", mode="window")
            bounds = snapshot["data"].get("bounds", {})
            wl = bounds.get("left", 0)
            wt = bounds.get("top", 0)
            ww = bounds.get("width", 1100)
            wh = bounds.get("height", 800)
            cx = wl + ww // 2
            cy = wt + wh - 70
            click_coordinates(int(cx), int(cy))
        except Exception as e:
            print(f"[3/4] Click fallback: {e}")
    time.sleep(0.3)

    # Step 4: 发送消息
    print("\n[4/4] Sending message...")
    ok = send_message_via_clipboard(message)
    if not ok:
        return False, "发送失败"

    return True, None


def send_message_to_current(message: str) -> tuple:
    """给当前已打开的聊天窗口发送消息（不需要指定联系人）"""
    print(f"[Send to current] Message: {message}")

    # Step 1: 恢复并聚焦窗口
    ok, info = restore_wechat_window()
    if not ok:
        return False, info

    hwnd = get_wechat_hwnd()
    if hwnd:
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    time.sleep(0.3)

    # Step 2: 点击输入框
    ok, _ = click_on_text("输入", snapshot_id="latest")
    if not ok:
        try:
            snapshot = peekaboo_see(title="微信", mode="window")
            bounds = snapshot["data"].get("bounds", {})
            wl = bounds.get("left", 0)
            wt = bounds.get("top", 0)
            ww = bounds.get("width", 1100)
            wh = bounds.get("height", 800)
            click_coordinates(int(wl + ww // 2), int(wt + wh - 70))
        except Exception:
            pass
    time.sleep(0.3)

    # Step 3: 发送
    ok = send_message_via_clipboard(message)
    if not ok:
        return False, "发送失败"

    return True, None


def get_wechat_status() -> dict:
    """获取微信状态"""
    hwnd = get_wechat_hwnd()
    if not hwnd:
        return {"status": "not_running"}

    return {
        "status": "running",
        "peekaboo_available": PEEKABOO_PATH is not None,
        "peekaboo_path": PEEKABOO_PATH,
    }


# ========== MCP 协议 ==========

TOOLS = [
    {
        "name": "wechat_get_status",
        "description": "获取微信状态（运行状态、窗口位置、Peekaboo是否可用）",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "wechat_send_message",
        "description": "给指定联系人发送消息（通过搜索流程打开聊天 + 剪贴板粘贴发送）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "contact": {"type": "string", "description": "联系人名称"},
                "message": {"type": "string", "description": "消息内容"},
            },
            "required": ["message"],
        },
    },
    {
        "name": "wechat_send_to_current",
        "description": "给当前已打开的聊天窗口发送消息（不需要指定联系人）",
        "inputSchema": {
            "type": "object",
            "properties": {
                "message": {"type": "string", "description": "消息内容"},
            },
            "required": ["message"],
        },
    },
]


async def handle_tool(name: str, arguments: dict) -> dict:
    """处理 MCP 工具调用"""
    if name == "wechat_get_status":
        result = get_wechat_status()
        return {"content": [{"type": "text", "text": json.dumps(result, ensure_ascii=False)}]}

    elif name == "wechat_send_message":
        message = arguments.get("message")
        contact = arguments.get("contact")

        if not message:
            return {"content": [{"type": "text", "text": "[ERROR] 需要指定 message"}]}
        if not contact:
            return {"content": [{"type": "text", "text": "[ERROR] 需要指定 contact"}]}

        success, err = send_message_to_contact(contact, message)
        if success:
            return {"content": [{"type": "text", "text": f"[OK] 消息已发送到 {contact}\n内容: {message}"}]}
        else:
            return {"content": [{"type": "text", "text": f"[ERROR] {err}"}]}

    elif name == "wechat_send_to_current":
        message = arguments.get("message")
        if not message:
            return {"content": [{"type": "text", "text": "[ERROR] 需要指定 message"}]}

        success, err = send_message_to_current(message)
        if success:
            return {"content": [{"type": "text", "text": f"[OK] 消息已发送\n内容: {message}"}]}
        else:
            return {"content": [{"type": "text", "text": f"[ERROR] {err}"}]}

    return {"content": [{"type": "text", "text": "未知命令"}]}


async def main():
    """MCP 主循环"""
    while True:
        try:
            line = sys.stdin.readline()
            if not line:
                break

            request = json.loads(line)

            if request.get("method") == "tools/list":
                response = {"jsonrpc": "2.0", "id": request.get("id"), "result": {"tools": TOOLS}}
                print(json.dumps(response), flush=True)

            elif request.get("method") == "tools/call":
                tool_name = request.get("name")
                tool_args = request.get("arguments", {})
                result = await handle_tool(tool_name, tool_args)
                response = {"jsonrpc": "2.0", "id": request.get("id"), "result": result}
                print(json.dumps(response), flush=True)

        except Exception as e:
            print(json.dumps({"jsonrpc": "2.0", "error": str(e)}), flush=True)


if __name__ == "__main__":
    asyncio.run(main())