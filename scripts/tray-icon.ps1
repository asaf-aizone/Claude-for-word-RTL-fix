# Claude for Office RTL - system tray status indicator.
#
# Standalone PowerShell script that shows a colored tray icon reflecting the
# connection status of the Node injector (inject.js).
#
# Communication is two-way via files in %TEMP%:
#   claude-word-rtl.status       - aggregate state, drives icon color.
#                                  Contents (one line):
#                                    CONNECTED        - injector attached to >=1 CDP target
#                                    DISCONNECTED     - no targets attached / injector exited
#                                    ERROR:<message>  - a fault was reported
#                                  Missing file is treated as DISCONNECTED.
#   claude-office-rtl.apps.json  - per-app state, drives the 3 status labels.
#                                  Contents: {"Word":"CONNECTED","Excel":"DISCONNECTED",...}
#                                  Missing file is treated as all DISCONNECTED.
#
# v0.2.0 expanded this from Word-only to Word + Excel + PowerPoint. The mutex,
# status file, PID file, lock file paths all keep the legacy "claude-word-rtl"
# prefix so a v0.2.0 tray launching during a v0.1.x reinstall window does not
# spawn duplicates and the prior PID files are still recognised.
#
# Zero npm dependencies. Uses System.Windows.Forms.NotifyIcon only.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not $PSScriptRoot) { $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
$StatusFile     = Join-Path $env:TEMP 'claude-word-rtl.status'
$AppsStatusFile = Join-Path $env:TEMP 'claude-office-rtl.apps.json'
$PidFile        = Join-Path $env:TEMP 'claude-word-rtl.pid'
$LockFile       = Join-Path $env:TEMP 'claude-word-rtl.lock'
$TrayPidFile    = Join-Path $env:TEMP 'claude-word-rtl.tray.pid'
$InstallDir     = Split-Path -Parent $PSScriptRoot  # parent of \scripts

# Singleton enforcement: a global mutex guarantees only one tray process
# exists per user session. Second launches exit immediately so the user
# never sees duplicate icons in the notification area.
# The mutex name keeps the legacy "ClaudeWordRtl" prefix so a v0.2.0 tray
# launched during a v0.1.x reinstall window correctly defers to the running
# v0.1.x tray (and vice versa) instead of spawning a duplicate icon.
$createdNew = $false
$script:TrayMutex = New-Object System.Threading.Mutex($true, 'Global\ClaudeWordRtlTrayMutex', [ref]$createdNew)
if (-not $createdNew) {
    # Another tray already owns the mutex. Exit quietly.
    exit 0
}

# Write our PID so uninstall.bat can target only the tray we started, and
# doctor.bat can confirm the tray is alive without a costly WMI query.
Set-Content -Path $TrayPidFile -Value $PID -Encoding ASCII -ErrorAction SilentlyContinue

# Freshness threshold: if the status file hasn't been touched in this many
# seconds and the injector process is gone, we treat the status as stale.
$StaleSeconds = 15

# Office app metadata table. Single source of truth for the tray's per-app
# loops. Keep in sync with lib/office-apps.js APPS array on the Node side.
# DocCollection is the COM property name used to enumerate open documents
# on each app, used by the Connect flow to remember which docs to reopen
# after relaunch via the wrapper. ProcessName is the executable basename
# without .EXE (matches Get-Process -Name semantics).
#
# OptIn = $true marks an app that requires explicit per-launch consent and
# is NOT wired into the generic auto-launch / generic Connect-item loops.
# Outlook is opt-in because silent CDP attach to mail content is a higher-
# grade exposure than to Word/Excel/PowerPoint document panels (see
# docs/OUTLOOK-EXPANSION-PLAN.md section 3, and probe/README.md "silent
# CDP attach" for the M0 evidence). The dedicated Outlook Connect item
# with its warning dialog (section 4.2 of the plan) is added in M1d.
# Until then the Outlook entry exists only so the tray's status label and
# apps.json reader recognise the app and surface its state.
$Apps = @(
    @{ Name = 'Word';       ProcessName = 'WINWORD';  ProgId = 'Word.Application';       Wrapper = 'word-wrapper.bat';       DocCollection = 'Documents'    }
    @{ Name = 'Excel';      ProcessName = 'EXCEL';    ProgId = 'Excel.Application';      Wrapper = 'excel-wrapper.bat';      DocCollection = 'Workbooks'    }
    @{ Name = 'PowerPoint'; ProcessName = 'POWERPNT'; ProgId = 'PowerPoint.Application'; Wrapper = 'powerpoint-wrapper.bat'; DocCollection = 'Presentations' }
    @{ Name = 'Outlook';    ProcessName = 'OUTLOOK';  ProgId = 'Outlook.Application';    Wrapper = 'outlook-wrapper.bat';    DocCollection = $null;            OptIn = $true }
)

# Win32 DestroyIcon P/Invoke so we can release GDI handles produced by
# Bitmap.GetHicon() without leaking across icon swaps.
Add-Type -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)] public static extern bool DestroyIcon(System.IntPtr hIcon);' -Name NativeMethods -Namespace ClaudeWordRtl -PassThru | Out-Null

# Build solid-color 16x16 bitmaps and convert to icons. Using Bitmap
# avoids shipping .ico files and keeps colors meaningful (green = live,
# red = dead, gray = starting).
$script:IconHandles = @()
# Icon design: 16x16, status-colored rounded square background, white "O"
# (for Office) in the center, tiny white RTL arrow in the corner. The "O"
# conveys "Office" (covering Word, Excel, PowerPoint), the arrow conveys
# "RTL direction fix", and the fill color conveys injector state.
function New-ColorIcon([System.Drawing.Color]$color) {
    $bmp = New-Object System.Drawing.Bitmap 16, 16
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    # Background: rounded square, status color
    $bg = New-Object System.Drawing.SolidBrush $color
    $rect = New-Object System.Drawing.Rectangle 0, 0, 16, 16
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $r = 3
    $path.AddArc($rect.X, $rect.Y, $r*2, $r*2, 180, 90)
    $path.AddArc($rect.Right - $r*2 - 1, $rect.Y, $r*2, $r*2, 270, 90)
    $path.AddArc($rect.Right - $r*2 - 1, $rect.Bottom - $r*2 - 1, $r*2, $r*2, 0, 90)
    $path.AddArc($rect.X, $rect.Bottom - $r*2 - 1, $r*2, $r*2, 90, 90)
    $path.CloseFigure()
    $g.FillPath($bg, $path)

    $white = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)

    # RTL arrow across the top strip of the icon (y=1..4), pointing LEFT
    # to signal right-to-left direction. Clearly visible at 16x16 because
    # it takes ~10 pixels of width.
    $arrowHead = @(
        (New-Object System.Drawing.Point 3, 3)   # tip (leftmost)
        (New-Object System.Drawing.Point 7, 0)   # top
        (New-Object System.Drawing.Point 7, 6)   # bottom
    )
    $g.FillPolygon($white, $arrowHead)
    # Arrow shaft from head to right edge
    $shaftPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 2
    $g.DrawLine($shaftPen, 7, 3, 14, 3)

    # White "O" in the bottom half of the icon (Office-wide branding)
    $font  = New-Object System.Drawing.Font ('Segoe UI', 9, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $textRect = New-Object System.Drawing.RectangleF 0, 6, 16, 10
    $g.DrawString('O', $font, $white, $textRect, $sf)

    $bg.Dispose(); $path.Dispose(); $font.Dispose()
    $white.Dispose(); $sf.Dispose(); $shaftPen.Dispose()
    $g.Dispose()
    $hIcon = $bmp.GetHicon()
    $bmp.Dispose()
    $script:IconHandles += $hIcon
    return [System.Drawing.Icon]::FromHandle($hIcon)
}

$iconGreen = New-ColorIcon ([System.Drawing.Color]::FromArgb(40, 180, 70))
$iconRed   = New-ColorIcon ([System.Drawing.Color]::FromArgb(210, 50, 50))
$iconGray  = New-ColorIcon ([System.Drawing.Color]::FromArgb(140, 140, 140))

$tray = New-Object System.Windows.Forms.NotifyIcon
$tray.Icon = $iconGray
$tray.Text = 'Claude for Office RTL - starting...'
$tray.Visible = $true

# Right-click context menu
$menu = New-Object System.Windows.Forms.ContextMenuStrip

# Connect flow is implemented as a Timer-driven state machine so the UI
# thread is never blocked waiting for an Office app to close or launch.
# Blocking the UI thread freezes the tray menu (the user sees it "stuck"
# on screen until the handler returns). Using timers keeps the tray
# responsive and lets us show progress / error dialogs without racing
# the menu.
#
# State lives in $script:ConnectState across timer ticks. Only ONE Connect
# flow runs at a time across all three apps; the App field records which
# app it targets (used for the wrapper path, COM enum, and dialog titles).
$script:ConnectState = @{
    Phase         = 'Idle'    # Idle | WaitingForClose | Launching
    App           = $null     # one of the $Apps entries when Phase != Idle
    DocsToReopen  = @()
    WaitedMs      = 0
    DocIndex      = 0
    CloseTimer    = $null
    LaunchTimer   = $null
    DocsTimer     = $null
}

# Auto-launch of the injector when an Office app is already up but the
# injector is gone. Recovery path: the user clicked Connect (so the app
# was launched via its wrapper, with the WebView2 debug flag set
# per-process), the injector was started by the wrapper but crashed or
# was killed while the app stayed up. Without this, the tray would show
# red until the user manually re-Connected. The tray is the natural place
# to own the relaunch because it is already polling every 2s and already
# knows both pieces of state (any Office app running, injector alive).
#
# Note: this path does NOT enable RTL on an Office app that was launched
# directly (taskbar, Recent files, double-click on a .docx/.xlsx/.pptx).
# Such an app has no debug surface to attach to. The user must use Connect
# to relaunch it through the wrapper.
#
# $AutoLaunchLastMs implements a soft cooldown so that if the injector
# fails to stay alive (missing node_modules, Node uninstalled, crash on
# startup), we do not spin up a relaunch storm every 2 seconds.
$script:AutoLaunchLastMs       = 0
$script:AutoLaunchCooldownMs   = 30000  # 30s - humane if the injector is crash-looping

function Stop-ConnectTimers {
    foreach ($n in @('CloseTimer', 'LaunchTimer', 'DocsTimer')) {
        $t = $script:ConnectState[$n]
        if ($t) { $t.Stop(); $t.Dispose(); $script:ConnectState[$n] = $null }
    }
    $script:ConnectState.Phase = 'Idle'
    $script:ConnectState.App   = $null
}

function Start-Launch-Phase {
    # Called after the target Office app has exited (or was never running).
    # Launches the per-app wrapper with the first queued doc; subsequent
    # docs are opened by DocsTimer with spacing so the app has time to
    # come up before each.
    $app = $script:ConnectState.App
    if (-not $app) { Stop-ConnectTimers; return }
    $wrapper = Join-Path $InstallDir $app.Wrapper
    $docs = $script:ConnectState.DocsToReopen
    $script:ConnectState.Phase = 'Launching'
    $script:ConnectState.DocIndex = 0

    if ($docs.Count -eq 0) {
        Start-Process -WindowStyle Hidden -FilePath 'cmd.exe' -ArgumentList '/c', "`"$wrapper`""
        Stop-ConnectTimers
        return
    }

    Start-Process -WindowStyle Hidden -FilePath 'cmd.exe' -ArgumentList '/c', "`"$wrapper`" `"$($docs[0])`""
    $script:ConnectState.DocIndex = 1
    if ($docs.Count -eq 1) {
        Stop-ConnectTimers
        return
    }

    # Space additional docs out so the Office app attaches each one to
    # the same running process. First extra waits 3s, subsequent 400ms.
    $docsTimer = New-Object System.Windows.Forms.Timer
    $docsTimer.Interval = 3000
    $docsTimer.add_Tick({
        $i = $script:ConnectState.DocIndex
        $docs = $script:ConnectState.DocsToReopen
        $a = $script:ConnectState.App
        if (-not $a) { Stop-ConnectTimers; return }
        $wrapper = Join-Path $InstallDir $a.Wrapper
        if ($i -ge $docs.Count) {
            Stop-ConnectTimers
            return
        }
        Start-Process -WindowStyle Hidden -FilePath 'cmd.exe' -ArgumentList '/c', "`"$wrapper`" `"$($docs[$i])`""
        $script:ConnectState.DocIndex = $i + 1
        $script:ConnectState.DocsTimer.Interval = 400
    })
    $script:ConnectState.DocsTimer = $docsTimer
    $docsTimer.Start()
}

function Start-ConnectFor($app) {
    # Per-app Connect handler. Identical state machine to the Word-only
    # v0.1.x flow, generalized over $app. The 3 menu items (Connect Word,
    # Connect Excel, Connect PowerPoint) all dispatch here.

    # Guard against double-clicks and against starting a second Connect
    # while a previous one (possibly for a different app) is mid-flight.
    if ($script:ConnectState.Phase -ne 'Idle') { return }

    $wrapper = Join-Path $InstallDir $app.Wrapper
    if (-not (Test-Path $wrapper)) { return }

    $running = Get-Process -Name $app.ProcessName -ErrorAction SilentlyContinue
    $docsToReopen = @()
    $hasUnsaved = $false

    if ($running) {
        # Enumerate open documents via COM before closing the app so we
        # can reopen them under the RTL session. Untitled new documents
        # have no real path - skip and warn. Each Office app exposes a
        # different collection (Documents/Workbooks/Presentations) with
        # the same .FullName item shape, so the loop is identical.
        try {
            $comApp = [Runtime.InteropServices.Marshal]::GetActiveObject($app.ProgId)
            $collection = $comApp.($app.DocCollection)
            foreach ($doc in $collection) {
                $full = $doc.FullName
                if ($full -and ($full -match '[\\/:]')) {
                    $docsToReopen += $full
                } else {
                    $hasUnsaved = $true
                }
            }
            [Runtime.InteropServices.Marshal]::ReleaseComObject($comApp) | Out-Null
        } catch {
            $docsToReopen = @()
        }

        $docLine = if ($docsToReopen.Count -gt 0) {
            "`nOpen documents will be reopened automatically:`n" + (($docsToReopen | ForEach-Object { '  - ' + (Split-Path -Leaf $_) }) -join "`n") + "`n"
        } else { '' }
        $unsavedLine = if ($hasUnsaved) {
            "`nWARNING: you have at least one UNSAVED document. Save it first, or it will be lost when $($app.Name) closes.`n"
        } else { '' }

        $ans = [System.Windows.Forms.MessageBox]::Show(
            "$($app.Name) is currently running without the RTL debug flag.`n`n" +
            "To enable the RTL fix, $($app.Name) must be closed and reopened." +
            $docLine + $unsavedLine + "`n" +
            "Close $($app.Name) now and relaunch with RTL?",
            "Claude for Office RTL - Connect $($app.Name)",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { return }

        # Kick off graceful close, then poll asynchronously via a Timer.
        # The handler returns immediately so the tray menu/UI stays live.
        $script:ConnectState.App = $app
        $script:ConnectState.DocsToReopen = $docsToReopen
        $script:ConnectState.WaitedMs = 0
        $script:ConnectState.Phase = 'WaitingForClose'
        $running | ForEach-Object { $_.CloseMainWindow() | Out-Null }

        $closeTimer = New-Object System.Windows.Forms.Timer
        $closeTimer.Interval = 250
        $closeTimer.add_Tick({
            $a = $script:ConnectState.App
            if (-not $a) { Stop-ConnectTimers; return }
            $stillRunning = [bool](Get-Process -Name $a.ProcessName -ErrorAction SilentlyContinue)
            if (-not $stillRunning) {
                $script:ConnectState.CloseTimer.Stop()
                $script:ConnectState.CloseTimer.Dispose()
                $script:ConnectState.CloseTimer = $null
                # Brief pause for the WebView2 debug port to free up
                # before relaunch. Even with dynamic ports we keep the
                # delay because the Office process itself needs a moment
                # to fully unwind before we spawn a new one.
                $delay = New-Object System.Windows.Forms.Timer
                $delay.Interval = 500
                $delay.add_Tick({
                    $script:ConnectState.LaunchTimer.Stop()
                    $script:ConnectState.LaunchTimer.Dispose()
                    $script:ConnectState.LaunchTimer = $null
                    Start-Launch-Phase
                })
                $script:ConnectState.LaunchTimer = $delay
                $delay.Start()
                return
            }
            $script:ConnectState.WaitedMs += 250
            if ($script:ConnectState.WaitedMs -ge 10000) {
                $script:ConnectState.CloseTimer.Stop()
                $script:ConnectState.CloseTimer.Dispose()
                $script:ConnectState.CloseTimer = $null
                $aName = $script:ConnectState.App.Name
                $force = [System.Windows.Forms.MessageBox]::Show(
                    "$aName did not close within 10 seconds.`n`n" +
                    "This usually means $aName is showing a dialog (save prompt, add-in message) that is blocking shutdown.`n`n" +
                    "Press OK to force-close $aName and relaunch with RTL. WARNING: any unsaved changes will be LOST.`n`n" +
                    "Press Cancel to leave $aName as-is. You can respond to the dialog and try Connect again.",
                    "Claude for Office RTL - Connect $aName",
                    [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($force -eq [System.Windows.Forms.DialogResult]::OK) {
                    Get-Process -Name $script:ConnectState.App.ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Milliseconds 400  # short, after force kill
                    Start-Launch-Phase
                } else {
                    Stop-ConnectTimers
                }
            }
        })
        $script:ConnectState.CloseTimer = $closeTimer
        $closeTimer.Start()
        return
    }

    # App not running - launch directly.
    $script:ConnectState.App = $app
    $script:ConnectState.DocsToReopen = @()
    Start-Launch-Phase
}

# --- Status labels (top of menu, refreshed each tick) ---
# 3 disabled menu items showing per-app state: "Word: connected", etc.
# Stored in a parallel array so the tick handler can update Text in place.
$script:StatusItems = @{}
foreach ($a in $Apps) {
    $mi = $menu.Items.Add("$($a.Name): not running")
    $mi.Enabled = $false
    $script:StatusItems[$a.Name] = $mi
}

$menu.Items.Add('-') | Out-Null  # separator

# --- Connect items, one per app ---
$script:ConnectItems = @{}
foreach ($a in $Apps) {
    # OptIn apps (Outlook) get a dedicated Connect item with a warning
    # dialog and an opt-in flag write, added in M1d. The generic Connect
    # flow (close-then-relaunch with no consent UX) is unsafe for them
    # because of the mail-content exposure model.
    if ($a.OptIn) { continue }
    $captured = $a   # capture for closure (PowerShell foreach var is shared)
    $mi = $menu.Items.Add("Connect $($captured.Name)")
    # ScriptBlock::Create + GetNewClosure avoids the classic PowerShell
    # foreach-closure-trap where every handler ends up bound to the last
    # iteration's $a. Using a parameterized scriptblock invoked with the
    # captured value sidesteps the trap cleanly.
    $handler = {
        param($appArg)
        Start-ConnectFor $appArg
    }.GetNewClosure()
    $mi.add_Click({
        Start-ConnectFor $captured
    }.GetNewClosure()) | Out-Null
    $script:ConnectItems[$a.Name] = $mi
}

$menu.Items.Add('-') | Out-Null  # separator

# --- Disconnect all ---
$miDisconnectAll = $menu.Items.Add('Disconnect all')
$miDisconnectAll.add_Click({
    # Tears down whatever state is active. Three independent things might
    # need stopping: a Connect flow mid-flight, the injector, and any of
    # the three Office apps. We stop each one if present. This makes
    # Disconnect-all the universal "recover from any state" button, which
    # is important because a failed Connect (app refused to launch,
    # injector running but cannot attach) previously left the user
    # stranded.

    # 1. Cancel any in-progress Connect flow so its Timers stop firing.
    if ($script:ConnectState -and $script:ConnectState.Phase -ne 'Idle') {
        Stop-ConnectTimers
    }

    # 2. Close every Office app we know about. WebView2 shuts down with
    #    the host app, which closes the debug port; inject.js keeps
    #    polling and flips DISCONNECTED within ~2 seconds.
    #    OptIn apps (Outlook) are skipped here: the user did not launch
    #    them through our Connect flow (Connect Outlook is M1d), so we
    #    must not close mail or calendar windows under their feet. The
    #    injector-side detach below is enough to drop the CDP attachment.
    $anyClosed = $false
    foreach ($a in $Apps) {
        if ($a.OptIn) { continue }
        $running = Get-Process -Name $a.ProcessName -ErrorAction SilentlyContinue
        if ($running) {
            $running | ForEach-Object { $_.CloseMainWindow() | Out-Null }
            $anyClosed = $true
        }
    }
    if ($anyClosed) {
        Start-Sleep -Milliseconds 800
        # Force-kill anything that refused to close gracefully.
        foreach ($a in $Apps) {
            if ($a.OptIn) { continue }
            Get-Process -Name $a.ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }

    # 3. If the injector is still alive after all apps went down - or if
    #    no app was running and only the injector needed stopping - kill
    #    it via its PID file. Handles the stuck state where a Connect
    #    failed partway: injector was launched but the host app never
    #    came up, leaving a live injector with no target.
    if (Test-Path $PidFile) {
        $pidLine = Get-Content -Path $PidFile -TotalCount 1 -ErrorAction SilentlyContinue
        $pidInt = if ($pidLine) { ("$pidLine").Trim() -as [int] } else { $null }
        if ($pidInt) {
            Stop-Process -Id $pidInt -Force -ErrorAction SilentlyContinue
        }
        Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $LockFile) { Remove-Item $LockFile -Force -ErrorAction SilentlyContinue }
    Set-Content -Path $StatusFile -Value 'DISCONNECTED' -Encoding ASCII -ErrorAction SilentlyContinue
    # Best-effort clear of the per-app status file too. The injector
    # would normally rewrite this on its next tick, but since we just
    # killed the injector we leave a clean slate for the tray's next read.
    try {
        $allOff = @{}
        foreach ($a in $Apps) { $allOff[$a.Name] = 'DISCONNECTED' }
        $tmp = $AppsStatusFile + '.tmp'
        ($allOff | ConvertTo-Json -Compress) | Set-Content -Path $tmp -Encoding ASCII -ErrorAction SilentlyContinue
        Move-Item -Path $tmp -Destination $AppsStatusFile -Force -ErrorAction SilentlyContinue
    } catch { }
}) | Out-Null

$menu.Items.Add('-') | Out-Null  # separator

$miShowLog = $menu.Items.Add('Show diagnostic log')
$miShowLog.add_Click({
    # Opens the injector's rolling log in notepad. Useful when a user
    # reports "Connect did not work" - the log shows which CDP targets
    # were discovered and whether attach succeeded. The file is
    # truncated at the start of every injector run, so it only contains
    # the current session's activity.
    $logPath = Join-Path $env:TEMP 'claude-word-rtl.log'
    if (-not (Test-Path $logPath)) {
        # Create empty file so notepad does not show the "create file?"
        # dialog. The injector will overwrite this when it next runs.
        Set-Content -Path $logPath -Value '' -Encoding ASCII -ErrorAction SilentlyContinue
    }
    Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$logPath`""
}) | Out-Null

$script:miCheckUpdate = $menu.Items.Add('Check for updates...')
$script:miCheckUpdate.add_Click({
    # Runs scripts/check-update.js and shows the result in a dialog. The
    # script emits one of three line prefixes: [UP TO DATE], [UPDATE
    # AVAILABLE] ... Download: <url>, or [ERROR] ... . Everything else
    # falls through to the generic Warning dialog with a fallback link.
    # The menu item is disabled during the call to prevent double-runs
    # from a stuck network (check-update.js has a 5s timeout).
    $script:miCheckUpdate.Enabled = $false
    try {
        $checkScript = Join-Path $InstallDir 'scripts\check-update.js'
        if (-not (Test-Path $checkScript)) {
            [System.Windows.Forms.MessageBox]::Show(
                "check-update.js was not found at:`n$checkScript",
                'Claude for Office RTL - Check for updates',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return
        }

        # 2>&1 merges stderr (used by [ERROR] lines) into stdout.
        $raw = & node $checkScript 2>&1 | Out-String
        $line = ($raw -split "`r?`n" |
                 Where-Object { $_.Trim() -ne '' } |
                 Select-Object -First 1)
        if (-not $line) { $line = '[ERROR] No output from check-update.js' }

        if ($line -match '^\[UP TO DATE\]\s*Local version:\s*(\S+)') {
            $local = $matches[1]
            [System.Windows.Forms.MessageBox]::Show(
                "Claude for Office RTL Fix is up to date (v$local).",
                'Claude for Office RTL - Check for updates',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        elseif ($line -match '^\[UPDATE AVAILABLE\]\s*Local\s+(\S+),\s*latest\s+(\S+)\.\s*Download:\s*(\S+)') {
            $local = $matches[1]; $latest = $matches[2]; $url = $matches[3]
            # Surface the install folder path so the user knows WHERE to
            # extract the new zip. Without this the user has to remember
            # where they originally installed (could be Downloads, Tools,
            # Desktop, anywhere) - common point of confusion in upgrades.
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "A newer version is available.`n`n" +
                "Installed: v$local`n" +
                "Latest:    v$latest`n`n" +
                "To upgrade:`n" +
                "  1. Download the zip (press OK below).`n" +
                "  2. Extract it OVER your install folder:`n" +
                "       $InstallDir`n" +
                "  3. Run install.bat again. It will stop the old tray`n" +
                "     and injector, then start the new ones.`n`n" +
                "Press OK to open the download page, or Cancel to skip.",
                'Claude for Office RTL - Check for updates',
                [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            if ($ans -eq [System.Windows.Forms.DialogResult]::OK) {
                Start-Process $url
                # Also open the install folder in Explorer so the user can
                # see where to drop the extracted files. Two windows in
                # the browser + in Explorer is the shortest path to a
                # successful upgrade.
                try { Start-Process explorer.exe -ArgumentList "`"$InstallDir`"" -ErrorAction SilentlyContinue } catch { }
            }
        }
        else {
            # [ERROR], network failure, 404 before first release, or any
            # unexpected output. Show the raw line plus a manual link.
            $fallbackUrl = 'https://github.com/asaf-aizone/Claude-for-Office-RTL-fix/releases/latest'
            [System.Windows.Forms.MessageBox]::Show(
                "Could not check for updates.`n`n" +
                $line.Trim() + "`n`n" +
                "You can check manually at:`n$fallbackUrl",
                'Claude for Office RTL - Check for updates',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to run check-update.js:`n$($_.Exception.Message)",
            'Claude for Office RTL - Check for updates',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $script:miCheckUpdate.Enabled = $true
    }
}) | Out-Null

$menu.Items.Add('-') | Out-Null  # separator

$miUninstall = $menu.Items.Add('Uninstall...')
$miUninstall.add_Click({
    # Confirm once, then launch uninstall.bat in a visible cmd window so the
    # user can see the cleanup progress. The tray itself must exit BEFORE
    # uninstall.bat runs, because uninstall.bat kills the tray as part of
    # its stop-tray-and-injector step - if we're still holding the mutex
    # when that happens, the hand-off is racy. Order: tray cleans up own
    # state, Start-Process detaches the cmd, then Application.Exit.
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Uninstall Claude for Office RTL Fix?`n`n" +
        "This will:`n" +
        "  - Remove the Startup entry and the Apps and Features registration`n" +
        "  - Stop the tray and the injector`n" +
        "  - Clean node_modules and temp status files`n" +
        "  - Remove any legacy WebView2 env var written by older installs`n`n" +
        "Word, Excel and PowerPoint themselves are not modified. The install" +
        " folder stays in place - you can delete it manually afterward.",
        'Claude for Office RTL - Uninstall',
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $uninstall = Join-Path $InstallDir 'uninstall.bat'
    if (-not (Test-Path $uninstall)) {
        [System.Windows.Forms.MessageBox]::Show(
            "uninstall.bat was not found next to the tray script. Cannot continue.",
            'Claude for Office RTL - Uninstall',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    # Clean up our own tray state first so uninstall.bat does not have to
    # kill us mid-flight.
    Stop-ConnectTimers
    $tray.Visible = $false
    $tray.Dispose()
    foreach ($h in $script:IconHandles) {
        [ClaudeWordRtl.NativeMethods]::DestroyIcon($h) | Out-Null
    }
    Remove-Item $TrayPidFile -Force -ErrorAction SilentlyContinue
    if ($script:TrayMutex) {
        try { $script:TrayMutex.ReleaseMutex() } catch {}
        $script:TrayMutex.Dispose()
    }

    # Launch uninstall.bat detached, then exit.
    Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', "`"$uninstall`""
    [System.Windows.Forms.Application]::Exit()
}) | Out-Null

$miExit = $menu.Items.Add('Exit')
$miExit.add_Click({
    $tray.Visible = $false
    $tray.Dispose()
    foreach ($h in $script:IconHandles) {
        [ClaudeWordRtl.NativeMethods]::DestroyIcon($h) | Out-Null
    }
    Remove-Item $TrayPidFile -Force -ErrorAction SilentlyContinue
    if ($script:TrayMutex) {
        try { $script:TrayMutex.ReleaseMutex() } catch {}
        $script:TrayMutex.Dispose()
    }
    [System.Windows.Forms.Application]::Exit()
}) | Out-Null

$tray.ContextMenuStrip = $menu

# Read the per-app status JSON written by inject.js. Returns a hashtable
# keyed by app name (Word, Excel, PowerPoint) with values 'CONNECTED',
# 'DISCONNECTED', or 'ERROR:<code>'. Missing/parse-error file is treated
# as all-DISCONNECTED so the tray degrades gracefully when the injector
# has not started yet (or has just been killed).
function Get-AppsStatus {
    $result = @{}
    foreach ($a in $Apps) { $result[$a.Name] = 'DISCONNECTED' }
    if (-not (Test-Path $AppsStatusFile)) { return $result }
    try {
        $raw = Get-Content -Path $AppsStatusFile -Raw -ErrorAction Stop
        if (-not $raw) { return $result }
        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
        foreach ($a in $Apps) {
            $val = $parsed.($a.Name)
            if ($val) { $result[$a.Name] = "$val" }
        }
    } catch {
        # parse error - leave the all-DISCONNECTED default
    }
    return $result
}

# Translate one app's raw apps.json value + process-running fact into the
# user-facing state string shown on the disabled status menu item.
#   not running         - no host process exists at all
#   connected           - process exists and apps.json says CONNECTED
#   running without RTL - process exists and apps.json says DISCONNECTED
#   error: <code>       - process exists and apps.json says ERROR:<code>
# The "running without RTL" string is what tells the user that this Connect
# button is the recovery path for a directly-launched (non-wrapper) Office app.
function Get-EffectiveAppState($app, $appsStatus) {
    $running = [bool](Get-Process -Name $app.ProcessName -ErrorAction SilentlyContinue)
    if (-not $running) { return 'not running' }
    $raw = $appsStatus[$app.Name]
    if (-not $raw) { return 'running without RTL' }
    if ($raw -eq 'CONNECTED') { return 'connected' }
    if ($raw -like 'ERROR:*') {
        $code = $raw.Substring(6)
        return "error: $code"
    }
    return 'running without RTL'
}

# Poll the status file every 2s (matches injector's POLL_MS) and update
# icon, tooltip, status labels, and per-item Enabled state.
#
# Aggregate icon color (unchanged from v0.1.x logic):
#   green = at least one app connected
#   red   = any error, or all disconnected with the injector having reported
#   gray  = startup, before any state has been observed
#
# Status labels (3 disabled items at top of menu) are refreshed every tick
# from apps.json + Get-Process. Connect items are enabled when their app
# is closed OR running without RTL; disabled when connected or while a
# Connect flow is in progress for ANY app. Disconnect-all is enabled when
# any app is alive, the injector is alive, or a Connect flow is running.
$script:lastStatus = ''
$tickAction = {
    $raw = 'DISCONNECTED'
    if (Test-Path $StatusFile) {
        try { $raw = (Get-Content -Path $StatusFile -TotalCount 1 -ErrorAction Stop).Trim() } catch { $raw = 'DISCONNECTED' }
    }

    # Staleness check
    $injectorAlive = $false
    if (Test-Path $PidFile) {
        $pidLine = Get-Content -Path $PidFile -TotalCount 1 -ErrorAction SilentlyContinue
        $pidInt = if ($pidLine) { ("$pidLine").Trim() -as [int] } else { $null }
        if ($pidInt) {
            $injectorAlive = [bool](Get-Process -Id $pidInt -ErrorAction SilentlyContinue)
        }
    }
    $effective = $raw
    if (-not $injectorAlive -and (Test-Path $StatusFile)) {
        $age = (Get-Date) - (Get-Item $StatusFile).LastWriteTime
        if ($age.TotalSeconds -gt $StaleSeconds) { $effective = 'DISCONNECTED' }
    }

    # Per-app state for the status labels and Connect-item enablement.
    $appsStatus = Get-AppsStatus
    $perApp = @{}   # name -> effective state string
    $anyAppRunning = $false
    foreach ($a in $Apps) {
        $state = Get-EffectiveAppState $a $appsStatus
        $perApp[$a.Name] = $state
        # An OptIn app being up does NOT justify auto-launching the injector.
        # The user might have started Outlook normally with no intent to use
        # the RTL fix this session, and the injector's blocklist would just
        # idle on the target anyway (wasting a node process). Only the
        # dedicated Connect Outlook flow (M1d) should bring the injector up
        # for an OptIn app.
        if ($state -ne 'not running' -and -not $a.OptIn) { $anyAppRunning = $true }
        # Refresh the disabled status label in place.
        $mi = $script:StatusItems[$a.Name]
        if ($mi) { $mi.Text = "$($a.Name): $state" }
    }

    $connectInProgress = ($script:ConnectState.Phase -ne 'Idle')

    # Auto-launch the injector if ANY of the three Office apps is up but
    # no injector is attached, and we are not already in the middle of a
    # Connect. Recovery path for the case where the injector crashed
    # while an app (started via Connect) stayed up. If the app was
    # started directly without Connect, it has no debug surface and the
    # relaunched injector just idles harmlessly.
    if ($anyAppRunning -and -not $injectorAlive -and -not $connectInProgress) {
        $nowMs = [Environment]::TickCount
        if (($nowMs - $script:AutoLaunchLastMs) -ge $script:AutoLaunchCooldownMs) {
            $autoLaunchVbs = Join-Path $InstallDir 'inject-hidden.vbs'
            if (Test-Path $autoLaunchVbs) {
                $script:AutoLaunchLastMs = $nowMs
                try {
                    Start-Process -FilePath 'wscript.exe' -ArgumentList "`"$autoLaunchVbs`"" -WindowStyle Hidden -ErrorAction Stop
                } catch { }
            }
        }
    }

    # Connect items: enabled only when the target app is NOT connected
    # AND no Connect flow is currently in flight (for any app). Same
    # gating as v0.1.x but applied per-app instead of just to Word.
    foreach ($a in $Apps) {
        $mi = $script:ConnectItems[$a.Name]
        if (-not $mi) { continue }
        $state = $perApp[$a.Name]
        $mi.Enabled = (($state -ne 'connected') -and (-not $connectInProgress))
    }
    # Disconnect-all: enabled when there is ANY state to tear down.
    $miDisconnectAll.Enabled = ($anyAppRunning -or $injectorAlive -or $connectInProgress)

    # Aggregate icon + tooltip. We compute a signature so we only repaint
    # when something actually changed, since GDI handle churn is the
    # historical source of icon-leak bugs in this script.
    $connectedCount = 0
    foreach ($k in $perApp.Keys) { if ($perApp[$k] -eq 'connected') { $connectedCount++ } }
    $errorState = ($effective -like 'ERROR:*')
    $totalApps = $Apps.Count

    $sig = "$effective|$connectedCount|$totalApps|$errorState"
    if ($sig -eq $script:lastStatus) { return }
    $script:lastStatus = $sig

    if ($connectedCount -gt 0 -and -not $errorState) {
        $tray.Icon = $iconGreen
        $tray.Text = "Claude for Office RTL - connected ($connectedCount of $totalApps)"
    } elseif ($errorState) {
        $tray.Icon = $iconRed
        $code = $effective.Substring(6)
        # Human-friendly tooltip mapping. Fall back to the raw code when
        # we have no explicit mapping, so future error codes still render.
        switch ($code) {
            'port-9222-taken-by-other-app' {
                $tray.Text = 'Claude for Office RTL - port 9222 used by another app. Run doctor.bat.'
            }
            default {
                $msg = $code
                if ($msg.Length -gt 55) { $msg = $msg.Substring(0, 55) + '...' }
                $tray.Text = "Claude for Office RTL - error: $msg"
            }
        }
    } else {
        $tray.Icon = $iconRed
        $tray.Text = 'Claude for Office RTL - disconnected'
    }
}
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 2000
$timer.add_Tick($tickAction)
$timer.Start()

# Initial poll so the icon reflects reality immediately instead of after 2s.
& $tickAction

[System.Windows.Forms.Application]::Run()
