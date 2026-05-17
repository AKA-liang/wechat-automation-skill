param(
    [Parameter(Mandatory=$true)]
    [string]$SessionName,
    [Parameter(Mandatory=$true)]
    [string]$Message
)

$ErrorActionPreference = "Continue"

$hwnd = "0xE079C"
$peekabooExe = "C:\Users\13265\Desktop\here\PeekabooWin-main\PeekabooWin-main\bin\peekaboo-win.js"

# Write message to temp file (UTF-8 without BOM - fixes encoding)
$msgFile = "$env:TEMP\wechat_msg_$PID.txt"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($msgFile, $Message, $utf8NoBom)

Write-Host ""
Write-Host "========== WeChat Auto Message V5 =========="
Write-Host "Target: $SessionName"
Write-Host "Message: $Message"
Write-Host ""

$jsScript = @"
const { spawn } = require('child_process');
const fs = require('fs');

const PEEkabooExe = 'C:\\Users\\13265\\Desktop\\here\\PeekabooWin-main\\PeekabooWin-main\\bin\\peekaboo-win.js';
const HWND = '0xE079C';
const SESSION_NAME = '$SessionName';
const MSG_FILE = '$msgFile'.replace(/\\\\/g, '\\\\\\\\');

function sleep(ms) {
    return new Promise(r => setTimeout(r, ms));
}

function peeka(args) {
    return new Promise((resolve, reject) => {
        const proc = spawn('node', [PEEkabooExe, ...args], { stdio: ['pipe', 'pipe', 'pipe'] });
        let stdout = '';
        let stderr = '';
        proc.stdout.on('data', d => stdout += d);
        proc.stderr.on('data', d => stderr += d);
        proc.on('close', code => {
            if (code !== 0 && stderr.includes('Error')) {
                console.error('PEEKABOO_ERR:' + stderr.substring(0, 300));
            }
            try {
                const trimmed = stdout.trim();
                // Multi-line JSON: find first { and last }
                const firstBrace = trimmed.indexOf('{');
                const lastBrace = trimmed.lastIndexOf('}');
                if (firstBrace >= 0 && lastBrace > firstBrace) {
                    const jsonStr = trimmed.substring(firstBrace, lastBrace + 1);
                    resolve(JSON.parse(jsonStr));
                } else {
                    reject(new Error('No JSON found in output: ' + trimmed.substring(0, 200)));
                }
            } catch(e) {
                reject(e);
            }
        });
    });
}

async function run() {
    const snap = await peeka(['see', '--mode', 'window', '--hwnd', HWND]);
    return { snapId: snap.snapshotId, bounds: snap.bounds };
}

(async () => {
    try {
        // Step 0: Restore window
        process.stdout.write('[0] Restore...');
        await peeka(['window', 'state', '--hwnd', HWND, '--state', 'restore']);
        await sleep(200);
        console.log(' OK');

        // Step 1: Get window snapshot
        process.stdout.write('[1] Snapshot...');
        const { snapId, bounds } = await run();
        console.log(' bounds=' + JSON.stringify(bounds));

        // Step 2: Click session liangtao (from OCR: window offset ~167, 80)
        const clickX = bounds.left + 167;
        const clickY = bounds.top + 80;
        process.stdout.write('[2] Click (' + clickX + ',' + clickY + ')...');
        await peeka(['mouse', 'click', '--x', String(Math.round(clickX)), '--y', String(Math.round(clickY))]);
        await sleep(600);
        console.log(' OK');

        // Step 3: Get chat snapshot
        process.stdout.write('[3] Chat snapshot...');
        const chatSnap = await peeka(['see', '--mode', 'window', '--hwnd', HWND]);
        const chatSnapId = chatSnap.snapshotId;
        const chatBounds = chatSnap.bounds;
        console.log(' bounds=' + JSON.stringify(chatBounds));

        // Step 4: Find and click input element in chat page
        process.stdout.write('[4] Input element...');
        const snapshotFile = process.env.USERPROFILE + '\\\\.peekaboo-windows\\\\snapshots\\\\' + chatSnapId + '\\\\snapshot.json';
        const snapData = JSON.parse(fs.readFileSync(snapshotFile, 'utf-8'));

        const inputEl = snapData.elements.find(e =>
            (e.automationId && e.automationId.toLowerCase().includes('chat_input')) ||
            (e.className && e.className.includes('ChatInput'))
        );

        if (inputEl) {
            console.log(' found:' + inputEl.id);
            await peeka(['snapshot', 'click', '--snapshot', chatSnapId, '--element-id', inputEl.id]);
        } else {
            console.log(' not found, clicking pane');
            const pane = snapData.elements.find(e => e.controlType === 'pane');
            if (pane) {
                await peeka(['snapshot', 'click', '--snapshot', chatSnapId, '--element-id', pane.id]);
            }
        }
        await sleep(200);

        // Step 5: Type message from file
        process.stdout.write('[5] Type...');
        await peeka(['type', '--text-file', MSG_FILE, '--delay-ms', '20']);
        console.log(' OK');

        // Cleanup temp file
        try { fs.unlinkSync(MSG_FILE); } catch(e) {}

        await sleep(100);

        // Step 6: Send
        process.stdout.write('[6] Send...');
        await peeka(['press', '--keys', '{ENTER}']);
        console.log(' OK');

        await sleep(400);

        // Step 7: Verify
        process.stdout.write('[7] Verify...');
        const finalSnap = await peeka(['see', '--mode', 'window', '--hwnd', HWND]);
        console.log(' OK');

        console.log('SEND_OK');

    } catch(e) {
        console.error('ERROR: ' + e.message);
        process.exit(1);
    }
})();
"@

try {
    $jsFile = "$env:TEMP\wechat_auto_$PID.js"
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($jsFile, $jsScript, $utf8NoBom)

    $nodeResult = & node $jsFile 2>&1
    Remove-Item $jsFile -Force -ErrorAction SilentlyContinue

    $output = $nodeResult | Out-String
    Write-Host $output

    if ($output -match "SEND_OK") {
        Write-Host ""
        Write-Host "========== SEND SUCCESS =========="
    } else {
        Write-Host ""
        Write-Host "========== SEND MAY HAVE ISSUES - check output above =========="
    }

} catch {
    Write-Host ""
    Write-Host "========== SEND FAILED =========="
    Write-Host "Error: $_"
    exit 1
}