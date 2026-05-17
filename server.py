# server.py
# 整合后的 WeChat MCP Server
# 核心思路：用 PeekabooWin OCR 文本点击（click --on）定位联系人，
#         用 PyAutoGUI + 剪贴板发送中文，绕过输入法问题

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
    click_coordinates,
    type_text,
    press_keys,
    focus_window,
    PEEKABOO_PATH,
)


# ========== 核心发送流程 ==========

def send_message_to_contact(contact_name: str, message: str) -> tuple:
    """给指定联系人发送消息（完整流程）"""
    print(f"[Send] Target: {contact_name}, Message: {message}")

    # Step 0: 恢复微信窗口
    print("\n[0/6] Restoring WeChat window...")
    ok, info = restore_wechat_window()
    if not ok:
        return False, info
    print(f"[0/6] Window restored: {info}")

    # Step 1: 聚焦窗口
    print("\n[1/6] Focusing window...")
    focus_window("微信")
    time.sleep(0.5)

    # Step 2: 用 OCR 文本点击找到并点击联系人
    print(f"\n[2/6] Clicking contact: {contact_name}...")
    ok, output = click_on_text(contact_name, snapshot_id="latest")
    if not ok:
        return False, f"Cannot find contact '{contact_name}' on screen: {output[:200]}"

    time.sleep(2)  # 等待聊天页加载

    # Step 3: 等待聊天页加载后再次截图确认
    print("\n[3/6] Verifying chat page...")
    try:
        snapshot = peekaboo_see(title="微信", mode="window")
        print(f"[3/6] Chat page verified, elements: {len(snapshot['data'].get('elements', []))}")
    except Exception as e:
        print(f"[3/6] Verification skip: {e}")

    # Step 4: 点击输入框（用 OCR 找"输入"相关文字，否则坐标降级）
    print("\n[4/6] Clicking input field...")
    _click_input_field()

    # Step 5: 输入消息（中文用剪贴板，英文用 peekaboo type）
    print("\n[5/6] Typing message...")
    _type_message(message)

    # Step 6: 发送
    print("\n[6/6] Pressing Enter to send...")
    press_keys("{ENTER}")
    time.sleep(0.5)

    return True, None


def _click_input_field():
    """点击输入框 - 优先用 OCR 找"输入"相关文字，否则坐标降级"""
    # 尝试用 OCR 点击输入框提示文字
    ok, _ = click_on_text("输入", snapshot_id="latest")
    if ok:
        print("[4/6] Input field clicked via OCR")
        return

    # 降级：坐标点击（窗口中央下方）
    try:
        snapshot = peekaboo_see(title="微信", mode="window")
        bounds = snapshot["data"].get("bounds", {})
        win_left = bounds.get("left", 0)
        win_top = bounds.get("top", 0)
        win_width = bounds.get("width", 1100)
        win_height = bounds.get("height", 800)
        input_x = win_left + win_width // 2
        input_y = win_top + win_height - 70
        click_coordinates(int(input_x), int(input_y))
        print(f"[4/6] Input field clicked via coords ({input_x}, {input_y})")
    except Exception as e:
        print(f"[4/6] Click fallback error: {e}")


def _type_message(message: str):
    """输入消息 - 优先 pyperclip 剪贴板（解决中文输入法问题）"""
    try:
        pyperclip.copy(message)
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.3)
    except Exception as e:
        # 降级：用 peekaboo type
        print(f"[5/6] Clipboard failed, using peekaboo type: {e}")
        type_text(message)


def send_message_to_current(message: str) -> tuple:
    """给当前已打开的聊天窗口发送消息（不需要指定联系人）"""
    print(f"[Send to current] Message: {message}")

    # Step 0: 恢复窗口
    ok, info = restore_wechat_window()
    if not ok:
        return False, info

    # Step 1: 聚焦窗口
    focus_window("微信")
    time.sleep(0.5)

    # Step 2: 找输入框并点击
    _click_input_field()
    time.sleep(0.3)

    # Step 3: 输入
    _type_message(message)

    # Step 4: 发送
    press_keys("{ENTER}")
    time.sleep(0.5)

    return True, None


def get_wechat_status() -> dict:
    """获取微信状态"""
    ok, info = restore_wechat_window()
    if not ok:
        return {"status": "not_running", "message": info}

    return {
        "status": "running",
        "peekaboo_available": PEEKABOO_PATH is not None,
        "peekaboo_path": PEEKABOO_PATH,
        "position": info["position"],
        "size": info["size"],
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
        "description": "给指定联系人发送消息（OCR文本点击联系人 + 剪贴板输入）",
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