# WeChat Automation Skill

Windows 微信自动化 Skill，支持向指定联系人发送消息。

## 功能

- 🔍 **联系人定位**：通过 PeekabooWin OCR 文本识别自动点击联系人
- 💬 **消息发送**：支持中文消息，使用剪贴板方式解决输入法兼容问题
- 🪟 **窗口恢复**：Win32 API 智能恢复被最小化/隐藏的微信窗口
- 🔌 **MCP 协议**：标准 JSON-RPC 2.0 接口，OpenClaw 可直接调用

## MCP 工具

| 工具 | 说明 |
|:---|:---|
| `wechat_get_status` | 获取微信状态（运行状态、窗口位置） |
| `wechat_send_message` | 给指定联系人发送消息 |
| `wechat_send_to_current` | 给当前已打开的聊天窗口发送消息 |

## 架构

```
window_restore.py       - Win32 API 窗口恢复（处理隐藏窗口）
peekaboo_integration.py - PeekabooWin CLI 封装 + OCR 文本点击
server.py              - MCP 主入口（stdin/stdout JSON-RPC）
```

## 依赖

- Python 3.10+
- [PeekabooWin](https://github.com/AKA-liang/PeekabooWin)（用于 UI 元素识别和 OCR）
- Python 包：`pyautogui`, `pygetwindow`, `pillow`, `pyperclip`, `opencv-python`, `pywin32`

## 快速使用

```bash
# 安装依赖（使用 conda 环境）
conda install -c conda-forge pyautogui pygetwindow pyperclip
pip install pillow opencv-python pywin32

# 或使用 pip 镜像
pip install pyautogui pygetwindow pillow pyperclip opencv-python pywin32 -i https://pypi.tuna.tsinghua.edu.cn/simple

# 运行测试
python test_integration.py

# 作为 MCP Server 运行
python server.py
```

## 工作流程

1. **窗口恢复**：Win32 `ShowWindow + MoveWindow` 确保微信窗口可见
2. **联系人定位**：PeekabooWin OCR 扫描 → `click --on "联系人名"` 精准点击
3. **输入框定位**：OCR 降级为窗口坐标计算
4. **消息输入**：`pyperclip.copy()` + `Ctrl+V` 粘贴（中文输入法兼容性）
5. **发送**：`Enter` 键发送

## 来源

本项目整合自两个方案：
- **dragon-wechat-mcp**（Python + pyautogui 坐标方案）
- **wechat-automation**（PeekabooWin PowerShell 方案）

保留两者的核心优势：PeekabooWin 的 UI 树精准定位 + Python MCP 的简洁协议 + 剪贴板中文输入

## License

MIT