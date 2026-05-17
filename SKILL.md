# WeChat Automation

Windows 微信自动化，基于 PeekabooWin OCR + Win32 API + Python MCP 协议整合方案。

## 触发词

微信发送、WeChat自动化、发微信

## MCP 工具

### wechat_get_status
获取微信运行状态和窗口信息。

### wechat_send_message
给指定联系人发送消息（自动定位联系人 + 剪贴板输入中文）。

```
参数：
- contact: 联系人名称（必填）
- message: 消息内容（必填）
```

### wechat_send_to_current
给当前已打开的聊天窗口发送消息（不指定联系人）。

```
参数：
- message: 消息内容（必填）
```

## 工作流程

```
1. 窗口最大化（Win32 SW_MAXIMIZE）→ 2. 聚焦微信窗口 → 3. Ctrl+F 激活搜索 → 4. 剪贴板粘贴联系人名 → 5. Enter 打开聊天 → 6. 点击输入框 → 7. 剪贴板粘贴消息 → 8. Enter 发送
```

## 技术要点

- **窗口固定**：每次操作前最大化微信窗口（坐标固定为 `-9,-9` 即屏幕左上角）
- **搜索流程**：用 `Ctrl+F` 激活微信内置搜索框，搜索结果第一条默认高亮，直接回车即打开聊天
- **中文输入**：使用 `pyperclip.copy()` + `Ctrl+V` 剪贴板粘贴（绕过输入法问题）
- **OCR 验证**：通过 PeekabooWin 截图 + OCR 识别聊天内联系人名出现次数（≥2 次视为打开成功）

## 依赖

- Python 3.10+（openclaw conda 环境）
- [PeekabooWin](https://github.com/AKA-liang/PeekabooWin)（安装在 `C:\Users\13265\Desktop\here\PeekabooWin-main\PeekabooWin-main`）
- Python 包：`pyautogui`, `pyperclip`, `pillow`, `pywin32`

## 注意事项

- 微信窗口需要保持打开状态
- 发送消息时会自动激活微信窗口
- 中文输入使用剪贴板方式（解决输入法兼容问题）
- PeekabooWin OCR 可识别聊天中的联系人名称，支持模糊匹配