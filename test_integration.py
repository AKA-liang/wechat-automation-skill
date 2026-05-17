# test_integration.py
# 整合方案测试脚本

import sys
import os
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from window_restore import restore_wechat_window
from peekaboo_integration import (
    peekaboo_see,
    click_on_text,
    PEEKABOO_PATH,
)
from server import send_message_to_current


def test_window_restore():
    print("=== Test: Window Restore ===")
    ok, info = restore_wechat_window()
    if ok:
        print(f"[OK] Window restored: {json.dumps(info, ensure_ascii=False, indent=2)}")
    else:
        print(f"[FAIL] {info}")
    print()


def test_peekaboo_see():
    print("=== Test: Peekaboo See ===")
    if not PEEKABOO_PATH:
        print("[SKIP] PeekabooWin not found")
        return

    try:
        snapshot = peekaboo_see(title="微信", mode="window")
        data = snapshot["data"]
        print(f"[OK] Snapshot ID: {snapshot['snapshotId']}")
        print(f"[OK] elementCount: {data.get('elementCount')}")
        print(f"[OK] bounds: {data.get('bounds')}")
        ocr_text = data.get("ocr", {}).get("text", "") if data.get("ocr") else ""
        print(f"[OK] ocr text (first 100): {ocr_text[:100]}")
    except Exception as e:
        print(f"[FAIL] {e}")
    print()


def test_send_message():
    print("=== Test: Send Message to Current ===")
    try:
        success, err = send_message_to_current("skill deploy test ok")
        if success:
            print("[OK] Message sent")
        else:
            print(f"[FAIL] {err}")
    except Exception as e:
        print(f"[FAIL] {e}")
    print()


if __name__ == "__main__":
    print("========== Integration Tests ==========\n")

    test_window_restore()
    test_peekaboo_see()
    test_send_message()

    print("========== Done ==========")