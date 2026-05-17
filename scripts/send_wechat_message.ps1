param(
    [Parameter(Mandatory=$true)]
    [string]$SessionName,
    [Parameter(Mandatory=$true)]
    [string]$Message,
    [string]$PeekabooPath = ""
)

$ErrorActionPreference = "Continue"

# ========== PeekabooWin Path Detection ==========
if (-not $PeekabooPath) {
    # Try development path first
    $devPath = "C:\Users\13265\Desktop\here\PeekabooWin-main\PeekabooWin-main"
    if (Test-Path (Join-Path $devPath "bin\peekaboo-win.js")) {
        $PeekabooPath = $devPath
        Write-Host "[PeekabooPath] Using dev path: $PeekabooPath"
    } else {
        # Try global npm
        $npmRoot = npm root -g 2>$null
        if ($npmRoot) {
            $globalPath = Join-Path $npmRoot "peekaboo-win"
            if (Test-Path (Join-Path $globalPath "bin\peekaboo-win.js")) {
                $PeekabooPath = $globalPath
                Write-Host "[PeekabooPath] Using global path: $PeekabooPath"
            }
        }
    }
}

if (-not $PeekabooPath -or -not (Test-Path (Join-Path $PeekabooPath "bin\peekaboo-win.js"))) {
    Write-Error "PeekabooWin not found. Please specify -PeekabooPath or install via: npm install -g peekaboo-win"
    exit 1
}

$peekabooExe = Join-Path $PeekabooPath "bin\peekaboo-win.js"
$nodeCmd = "node `"$peekabooExe`""

# ========== Win32 API for Window Restore ==========
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32Restore {
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    public struct RECT { public int Left, Top, Right, Bottom; }
    public const int SW_RESTORE = 9;
}
"@

function Restore-WeChatWindow {
    Write-Host "[Restore] Searching for WeChat window..."

    # Try FindWindow with WeChat title
    $hwnd = [Win32Restore]::FindWindow("Qt51514QWindowIcon", "微信")

    if ($hwnd -eq [IntPtr]::Zero) {
        # Fallback: use peekaboo to list windows and find WeChat hwnd
        Write-Host "[Restore] Trying via peekaboo windows list..."
        $output = Invoke-Expression "$nodeCmd windows list" 2>&1
        if ($output -match '"hwnd":\s*"0x([0-9A-Fa-f]+)".*?"title":\s*"微信"') {
            $hwndInt = [IntPtr][long]::Parse($Matches[1], 'HexNumber')
            $hwnd = $hwndInt
        }
    }

    if ($hwnd -ne [IntPtr]::Zero) {
        Write-Host "[Restore] Found WeChat window, hwnd=$($hwnd.ToInt64().ToString('X'))"

        # Get current window bounds
        $rect = New-Object Win32Restore+RECT
        [Win32Restore]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
        $width = $rect.Right - $rect.Left
        $height = $rect.Bottom - $rect.Top
        Write-Host "[Restore] Current bounds: ($($rect.Left), $($rect.Top)) size ${width}x${height}"

        # Check if window is off-screen or has abnormal size
        $offScreen = $rect.Left -lt -10000 -or $rect.Top -lt -10000
        $abnormalSize = $width -lt 200 -or $height -lt 200

        if ($offScreen -or $abnormalSize) {
            Write-Host "[Restore] Window is off-screen or has abnormal size, restoring..."

            # First SW_RESTORE to restore from minimized state
            [Win32Restore]::ShowWindow($hwnd, [Win32Restore]::SW_RESTORE) | Out-Null
            Start-Sleep -Milliseconds 300

            # Force move to visible area with normal size (1118x809 is known good size)
            $normalWidth = 1118
            $normalHeight = 809
            [Win32Restore]::MoveWindow($hwnd, 100, 100, $normalWidth, $normalHeight, $true) | Out-Null
            Start-Sleep -Milliseconds 500

            # Verify new bounds
            [Win32Restore]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
            Write-Host "[Restore] New bounds: ($($rect.Left), $($rect.Top)) size $($rect.Right - $rect.Left)x$($rect.Bottom - $rect.Top)"
        } else {
            # Normal visible window, just bring to foreground
            [Win32Restore]::ShowWindow($hwnd, [Win32Restore]::SW_RESTORE) | Out-Null
            Start-Sleep -Milliseconds 300
        }

        [Win32Restore]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 500
        Write-Host "[Restore] Window restored to foreground"
        return $true
    } else {
        Write-Warning "[Restore] WeChat window not found"
        return $false
    }
}

function Invoke-Peekaboo {
    param([string]$Args, [int]$Timeout = 30)
    $result = Invoke-Expression "$nodeCmd $Args" 2>&1
    if ($LASTEXITCODE -ne 0 -and ($result -match "Error" -or $result -match "exception")) {
        Write-Warning "Peekaboo warning: $($result | Out-String)"
    }
    return $result
}

function Get-SessionCoordinates {
    param([string]$Name)

    Write-Host "[Find] Searching for session: $Name"

    # Try peekaboo see
    $snapshotOutput = Invoke-Peekaboo "see --mode window --title `"微信`""
    if ($snapshotOutput -notmatch '"snapshotId":\s*"([^"]+)"') {
        throw "Cannot get snapshot ID from peekaboo"
    }
    $snapshotId = $Matches[1]
    Write-Host "[Find] Snapshot: $snapshotId"

    # Read snapshot JSON
    $snapshotFile = Join-Path $env:USERPROFILE ".peekaboo-windows\snapshots\$snapshotId\snapshot.json"
    if (-not (Test-Path $snapshotFile)) {
        throw "Snapshot file not found: $snapshotFile"
    }

    $snapshotData = Get-Content $snapshotFile -Raw | ConvertFrom-Json

    # Get window offset from snapshot bounds
    $windowLeft = 0
    $windowTop = 0
    if ($snapshotData.elements.Count -gt 0) {
        $firstElem = $snapshotData.elements[0]
        if ($firstElem.bounds) {
            $windowLeft = $firstElem.bounds.left
            $windowTop = $firstElem.bounds.top
        }
    }

    Write-Host "[Find] Window offset: left=$windowLeft, top=$windowTop"

    # Search for session name in elements
    foreach ($elem in $snapshotData.elements) {
        $elemName = $elem.name -replace "`n.*", ""
        if ($elemName -eq $Name) {
            $centerX = $elem.center.x
            $centerY = $elem.center.y
            $absX = $windowLeft + $centerX
            $absY = $windowTop + $centerY
            Write-Host "[Find] Session found at: ($absX, $absY)"
            return @{ X = $absX; Y = $absY; SnapshotId = $snapshotId }
        }
    }

    throw "Session not found: $Name"
}

function Get-ChatInputElement {
    Write-Host "[Input] Finding chat input field..."

    $snapshotOutput = Invoke-Peekaboo "see --mode window --title `"微信`""
    if ($snapshotOutput -notmatch '"snapshotId":\s*"([^"]+)"') {
        throw "Cannot get snapshot ID"
    }
    $snapshotId = $Matches[1]

    $snapshotFile = Join-Path $env:USERPROFILE ".peekaboo-windows\snapshots\$snapshotId\snapshot.json"
    $snapshotData = Get-Content $snapshotFile -Raw | ConvertFrom-Json

    foreach ($elem in $snapshotData.elements) {
        if ($elem.className -eq "mmui::ChatInputField") {
            Write-Host "[Input] Found: $($elem.id)"
            return @{ Id = $elem.id; SnapshotId = $snapshotId }
        }
    }

    throw "Chat input field not found"
}

# ========== Main Flow ==========

Write-Host ""
Write-Host "========== WeChat Automation - Send Message =========="
Write-Host "Target: $SessionName"
Write-Host "Message: $Message"
Write-Host ""

try {
    # Step 0: Restore WeChat window (critical if minimized/hidden)
    Write-Host "[0/6] Restoring WeChat window..."
    $restored = Restore-WeChatWindow
    if (-not $restored) {
        Write-Warning "[0/6] Could not restore window, trying anyway..."
    }

    # Step 1: Focus window
    Write-Host "[1/6] Focusing WeChat window..."
    Invoke-Peekaboo "window focus --title `"微信`"" | Out-Null
    Start-Sleep -Milliseconds 500

    # Step 2: Find session coordinates
    Write-Host "[2/6] Finding session coordinates..."
    $coords = Get-SessionCoordinates -Name $SessionName

    # Step 3: Click session
    Write-Host "[3/6] Clicking session at ($($coords.X), $($coords.Y))..."
    Invoke-Peekaboo "mouse click --x $($coords.X) --y $($coords.Y)" | Out-Null

    # Step 4: Wait for chat page
    Write-Host "[4/6] Waiting for chat page to load..."
    Start-Sleep -Seconds 2

    # Verify chat page by checking title
    $verifyOutput = Invoke-Peekaboo "see --mode window --title `"微信`""
    if ($verifyOutput -notmatch [regex]::Escape($SessionName)) {
        Write-Warning "[4/6] Chat page may not have loaded (session name not found in snapshot)"
    } else {
        Write-Host "[4/6] Chat page confirmed"
    }

    # Step 5: Find and click input field
    Write-Host "[5/6] Clicking chat input field..."
    $inputElem = Get-ChatInputElement
    $clickResult = Invoke-Peekaboo "snapshot click --snapshot $($inputElem.SnapshotId) --element-id $($inputElem.Id)" 2>&1
    Write-Host "[5/6] Click result: $($clickResult | Out-String)"
    Start-Sleep -Milliseconds 300

    # Step 6: Type message
    Write-Host "[6/6] Typing message: $Message"
    Invoke-Peekaboo "type --text `"$Message`"" | Out-Null
    Start-Sleep -Milliseconds 200

    # Step 7: Press Enter to send
    Write-Host "[7/6] Sending message..."
    Invoke-Peekaboo "press --keys `"{ENTER}`"" | Out-Null

    # Step 8: Verify
    Write-Host "[8/8] Verifying message..."
    Start-Sleep -Seconds 2
    $finalOutput = Invoke-Peekaboo "see --mode window --title `"微信`""

    if ($finalOutput -match [regex]::Escape($Message)) {
        Write-Host ""
        Write-Host "========== SEND SUCCESS =========="
        Write-Host "Message '$Message' sent to '$SessionName'"
        exit 0
    } else {
        Write-Warning "[Verify] Message not found in final snapshot, but may have been sent"
        Write-Host ""
        Write-Host "========== SEND COMPLETED (unverified) =========="
        exit 0
    }

} catch {
    Write-Host ""
    Write-Host "========== SEND FAILED =========="
    Write-Host "Error: $_"
    Write-Host $_.ScriptStackTrace
    exit 1
}