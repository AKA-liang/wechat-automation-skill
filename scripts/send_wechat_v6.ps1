param(
    [Parameter(Mandatory=$true)]
    [string]$SessionName,
    [Parameter(Mandatory=$true)]
    [string]$Message
)

$ErrorActionPreference = "Continue"

$peekabooExe = "C:\Users\13265\Desktop\here\PeekabooWin-main\PeekabooWin-main\bin\peekaboo-win.js"
$hwnd = "0xE079C"
$dragonPath = "C:\here\dragon-wechat-mcp-1.0.0"

# Write message to temp file (UTF-8 without BOM)
$msgFile = "$env:TEMP\wechat_msg_$PID.txt"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($msgFile, $Message, $utf8NoBom)

Write-Host ""
Write-Host "========== WeChat Hybrid V6 =========="
Write-Host "Target: $SessionName"
Write-Host "Message: $Message"
Write-Host ""

# ========== Helper: Run peekaboo, get string, parse JSON ==========
function Get-PeekabooJson {
    param([string[]]$Args)
    $result = & node $peekabooExe $Args 2>&1
    $str = "$result"
    $trimmed = $str.Trim()
    $first = $trimmed.IndexOf('{')
    $last = $trimmed.LastIndexOf('}')
    if ($first -ge 0 -and $last -gt $first) {
        $jsonStr = $trimmed.Substring($first, $last - $first + 1)
        try { return $jsonStr | ConvertFrom-Json } catch { }
    }
    return $null
}

# ========== Part 1: PeekabooWin (Steps 0-3) ==========
Write-Host "[P1] PeekabooWin: restore + find session + open chat"

# Step 0: Restore window
Write-Host "[P1-0] Restore window..."
Get-PeekabooJson @("window", "state", "--hwnd", $hwnd, "--state", "restore") | Out-Null
Start-Sleep -Milliseconds 400

# Step 1: Get window snapshot
Write-Host "[P1-1] Get snapshot..."
$snap = Get-PeekabooJson @("see", "--mode", "window", "--hwnd", $hwnd)
if (!$snap) {
    Write-Host "[P1-1] Failed to get snapshot"
    exit 1
}
$bounds = $snap.bounds
$snapId = $snap.snapshotId
Write-Host "[P1-1] snapshotId=$snapId, bounds=left=$($bounds.left), top=$($bounds.top)"

# Step 2: Find session in snapshot elements and click it
Write-Host "[P1-2] Finding session '$SessionName' in snapshot..."

$snapshotFile = "$env:USERPROFILE\.peekaboo-windows\snapshots\$snapId\snapshot.json"
if (!(Test-Path $snapshotFile)) {
    Write-Host "[P1-2] Snapshot file not found: $snapshotFile"
    exit 1
}

$snapData = Get-Content $snapshotFile -Raw | ConvertFrom-Json

# Get window offset from element bounds
$elem = $snapData.elements | Select-Object -First 1
$winLeft = if ($elem.bounds.left) { $elem.bounds.left } else { $bounds.left }
$winTop = if ($elem.bounds.top) { $elem.bounds.top } else { $bounds.top }
Write-Host "[P1-2] Window offset: left=$winLeft, top=$winTop"

$found = $false
foreach ($e in $snapData.elements) {
    $name = $e.name -replace "`n.*", ""
    if ($name -eq $SessionName) {
        $clickX = $winLeft + $e.center.x
        $clickY = $winTop + $e.center.y
        Write-Host "[P1-2] Found '$name' at center=($($e.center.x), $($e.center.y)) → abs=($clickX, $clickY)"
        Get-PeekabooJson @("mouse", "click", "--x", [Math]::Round($clickX), "--y", [Math]::Round($clickY)) | Out-Null
        $found = $true
        break
    }
}

if (!$found) {
    Write-Host "[P1-2] Exact match not found, trying partial..."
    foreach ($e in $snapData.elements) {
        $name = $e.name -replace "`n.*", ""
        if ($name -and $name.Contains($SessionName)) {
            $clickX = $winLeft + $e.center.x
            $clickY = $winTop + $e.center.y
            Write-Host "[P1-2] Partial match '$name' at ($clickX, $clickY)"
            Get-PeekabooJson @("mouse", "click", "--x", [Math]::Round($clickX), "--y", [Math]::Round($clickY)) | Out-Null
            $found = $true
            break
        }
    }
}

if (!$found) {
    Write-Host "[P1-2] Session not found in snapshot elements"
    exit 1
}

# Step 3: Wait for chat page
Write-Host "[P1-3] Waiting for chat page..."
Start-Sleep -Milliseconds 1000

Write-Host "[P1] Done - chat should be open"
Write-Host ""

# ========== Part 2: Python (Steps 4-6) ==========
Write-Host "[P2] Python: click input + type via clipboard + send"

$pythonScript = @"
import sys
import os
import time
sys.path.insert(0, r'$($dragonPath.Replace('\', '\\\\'))')

import pyautogui
pyautogui.FAILSAFE = False
import pyperclip
import pygetwindow as gw

def send_via_clipboard(message):
    wins = gw.getWindowsWithTitle('微信')
    win = None
    for w in wins:
        if w.width > 500 and w.width < 2000:
            win = w
            break

    if not win:
        print('[P2-ERR] No WeChat window found')
        return False

    # Click input box: center-bottom of window, 60px from bottom edge
    input_x = win.left + win.width // 2
    input_y = win.top + win.height - 60
    print(f'[P2-1] Clicking input at ({input_x}, {input_y})')
    pyautogui.click(input_x, input_y)
    time.sleep(0.3)

    # Copy to clipboard (avoids CLI encoding issues!)
    pyperclip.copy(message)
    time.sleep(0.1)

    # Paste
    print('[P2-2] Pasting from clipboard...')
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.4)

    # Send
    print('[P2-3] Pressing Enter...')
    pyautogui.press('enter')
    time.sleep(0.5)

    print('[P2] Done!')
    return True

if __name__ == '__main__':
    msg_file = r'$($msgFile.Replace('\', '\\\\'))'
    msg = open(msg_file, 'r', encoding='utf-8').read()
    print(f'[P2] Message: {repr(msg)}')
    success = send_via_clipboard(msg)
    try:
        os.unlink(msg_file)
    except:
        pass
    sys.exit(0 if success else 1)
"@

$pyFile = "$env:TEMP\wechat_send_$PID.py"
[System.IO.File]::WriteAllText($pyFile, $pythonScript, $utf8NoBom)

$pyResult = python $pyFile 2>&1 | Out-String
Write-Host $pyResult

Remove-Item $pyFile -Force -ErrorAction SilentlyContinue

if ($pyResult -match "Done!" -or $pyResult -match "P2.*success") {
    Write-Host ""
    Write-Host "========== SEND SUCCESS =========="
} else {
    Write-Host ""
    Write-Host "========== CHECK RESULT ABOVE =========="
}