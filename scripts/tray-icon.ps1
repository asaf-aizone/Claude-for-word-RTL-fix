# Claude for Word RTL - system tray status indicator.
#
# Standalone PowerShell script that shows a colored tray icon reflecting the
# connection status of the Node injector (inject.js).
#
# Communication is one-way via a status file:
#   %TEMP%\claude-word-rtl.status
# Contents (one line):
#   CONNECTED        - injector attached to at least one CDP target
#   DISCONNECTED     - no targets attached / injector exited
#   ERROR:<message>  - a fault was reported (e.g. DOM selectors no longer match)
# Missing file is treated as DISCONNECTED.
#
# Zero npm dependencies. Uses System.Windows.Forms.NotifyIcon only.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not $PSScriptRoot) { $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
$StatusFile = Join-Path $env:TEMP 'claude-word-rtl.status'
$PidFile    = Join-Path $env:TEMP 'claude-word-rtl.pid'
$LockFile   = Join-Path $env:TEMP 'claude-word-rtl.lock'
$TrayPidFile = Join-Path $env:TEMP 'claude-word-rtl.tray.pid'
$InstallDir = Split-Path -Parent $PSScriptRoot  # parent of \scripts

# Singleton enforcement: a global mutex guarantees only one tray process
# exists per user session. Second launches exit immediately so the user
# never sees duplicate icons in the notification area.
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

# Win32 DestroyIcon P/Invoke so we can release GDI handles produced by
# Bitmap.GetHicon() without leaking across icon swaps.
Add-Type -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)] public static extern bool DestroyIcon(System.IntPtr hIcon);' -Name NativeMethods -Namespace ClaudeWordRtl -PassThru | Out-Null

# Build two solid-color 16x16 bitmaps and convert to icons. Using Bitmap
# avoids shipping .ico files and keeps colors meaningful (green = live,
# red = dead, gray = starting).
$script:IconHandles = @()
# Icon design: 16x16, status-colored rounded square background, white "W"
# (for Word) in the center, tiny white RTL arrow in the corner. The W
# conveys "Word", the arrow conveys "RTL direction fix", and the fill
# color conveys injector state.
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

    # White "W" in the bottom half of the icon
    $font  = New-Object System.Drawing.Font ('Segoe UI', 9, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $textRect = New-Object System.Drawing.RectangleF 0, 6, 16, 10
    $g.DrawString('W', $font, $white, $textRect, $sf)

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
$tray.Text = 'Claude for Word RTL - starting...'
$tray.Visible = $true

# Right-click context menu
$menu = New-Object System.Windows.Forms.ContextMenuStrip

# Connect flow is implemented as a Timer-driven state machine so the UI
# thread is never blocked waiting for Word to close or launch. Blocking
# the UI thread freezes the tray menu (the user sees it "stuck" on screen
# until the handler returns). Using timers keeps the tray responsive and
# lets us show progress / error dialogs without racing the menu.
#
# State lives in $script:ConnectState across timer ticks.
$script:ConnectState = @{
    Phase         = 'Idle'    # Idle | WaitingForClose | Launching
    DocsToReopen  = @()
    WaitedMs      = 0
    DocIndex      = 0
    CloseTimer    = $null
    LaunchTimer   = $null
    DocsTimer     = $null
}

# Auto-launch of the injector when Word is already up but the injector
# is gone. Recovery path: the user clicked Connect (so Word was launched
# via word-wrapper.bat with the WebView2 debug flag set per-process),
# the injector was started by the wrapper but crashed or was killed
# while Word stayed up. Without this, the tray would show red until the
# user manually re-Connected.
#
# Note: this path does NOT enable RTL on a Word that was launched
# directly (taskbar, Recent files, double-click on a .docx). Such a
# Word has no debug surface to attach to. The user must use Connect
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
}

function Start-Launch-Phase {
    # Called after Word has exited (or was never running). Launches the
    # wrapper with the first queued doc; subsequent docs are opened by
    # DocsTimer with spacing so Word has time to come up before each.
    $wrapper = Join-Path $InstallDir 'word-wrapper.bat'
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

    # Space additional docs out so Word attaches each one to the same
    # running Winword. First extra waits 3s, subsequent 400ms.
    $docsTimer = New-Object System.Windows.Forms.Timer
    $docsTimer.Interval = 3000
    $docsTimer.add_Tick({
        $i = $script:ConnectState.DocIndex
        $docs = $script:ConnectState.DocsToReopen
        $wrapper = Join-Path $InstallDir 'word-wrapper.bat'
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

$miConnect = $menu.Items.Add('Connect (relaunch Claude for Word RTL Fix)')
$miConnect.add_Click({
    # Guard against double-clicks while a previous Connect is mid-flight.
    if ($script:ConnectState.Phase -ne 'Idle') { return }

    $wrapper = Join-Path $InstallDir 'word-wrapper.bat'
    if (-not (Test-Path $wrapper)) { return }

    $running = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    $docsToReopen = @()
    $hasUnsaved = $false

    if ($running) {
        # Enumerate open documents via COM before closing Word so we can
        # reopen them under the RTL session. Untitled new documents
        # (Document1, etc.) have no real path - skip and warn.
        try {
            $wordApp = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            foreach ($doc in $wordApp.Documents) {
                $full = $doc.FullName
                if ($full -and ($full -match '[\\/:]')) {
                    $docsToReopen += $full
                } else {
                    $hasUnsaved = $true
                }
            }
            [Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        } catch {
            $docsToReopen = @()
        }

        $docLine = if ($docsToReopen.Count -gt 0) {
            "`nOpen documents will be reopened automatically:`n" + (($docsToReopen | ForEach-Object { '  - ' + (Split-Path -Leaf $_) }) -join "`n") + "`n"
        } else { '' }
        $unsavedLine = if ($hasUnsaved) {
            "`nWARNING: you have at least one UNSAVED document. Save it first, or it will be lost when Word closes.`n"
        } else { '' }

        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Word is currently running without the RTL debug flag.`n`n" +
            "To enable the RTL fix, Word must be closed and reopened." +
            $docLine + $unsavedLine + "`n" +
            'Close Word now and relaunch with RTL?',
            'Claude for Word RTL - Connect',
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { return }

        # Kick off graceful close, then poll asynchronously via a Timer.
        # The handler returns immediately so the tray menu/UI stays live.
        $script:ConnectState.DocsToReopen = $docsToReopen
        $script:ConnectState.WaitedMs = 0
        $script:ConnectState.Phase = 'WaitingForClose'
        $running | ForEach-Object { $_.CloseMainWindow() | Out-Null }

        $closeTimer = New-Object System.Windows.Forms.Timer
        $closeTimer.Interval = 250
        $closeTimer.add_Tick({
            $stillRunning = [bool](Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
            if (-not $stillRunning) {
                $script:ConnectState.CloseTimer.Stop()
                $script:ConnectState.CloseTimer.Dispose()
                $script:ConnectState.CloseTimer = $null
                # Brief pause for port 9222 to free up before relaunch.
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
                $force = [System.Windows.Forms.MessageBox]::Show(
                    "Word did not close within 10 seconds.`n`n" +
                    "This usually means Word is showing a dialog (save prompt, add-in message) that is blocking shutdown.`n`n" +
                    "Press OK to force-close Word and relaunch with RTL. WARNING: any unsaved changes will be LOST.`n`n" +
                    "Press Cancel to leave Word as-is. You can respond to the dialog and try Connect again.",
                    'Claude for Word RTL - Connect',
                    [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($force -eq [System.Windows.Forms.DialogResult]::OK) {
                    Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
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

    # Word not running - launch directly.
    $script:ConnectState.DocsToReopen = @()
    Start-Launch-Phase
}) | Out-Null

$miDisconnect = $menu.Items.Add('Disconnect (close Claude for Word RTL Fix)')
$miDisconnect.add_Click({
    # Tears down whatever state is active. Three independent things might
    # need stopping: a Connect flow mid-flight, the injector, and Word
    # itself. We stop each one if present. This makes Disconnect the
    # universal "recover from any state" button, which is important because
    # a failed Connect (Word refused to launch, injector running but can
    # not attach) previously left the user stranded.

    # 1. Cancel any in-progress Connect flow so its Timers stop firing.
    if ($script:ConnectState -and $script:ConnectState.Phase -ne 'Idle') {
        Stop-ConnectTimers
    }

    # 2. Close Word if running. WebView2 shuts down with Word, which
    #    closes the debug port; inject.js keeps polling and flips to
    #    DISCONNECTED within ~2 seconds.
    $running = Get-Process -Name WINWORD -ErrorAction SilentlyContinue
    if ($running) {
        $running | ForEach-Object { $_.CloseMainWindow() | Out-Null }
        Start-Sleep -Milliseconds 800
        # Force-kill anything that refused to close gracefully.
        Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }

    # 3. If the injector is still alive after Word went down - or if Word
    #    was never running and only the injector needed stopping - kill it
    #    via its PID file. This handles the stuck state where a
    #    Connect failed partway: injector was launched by word-wrapper but
    #    Word itself never came up, leaving a live injector with no target.
    if (Test-Path $PidFile) {
        $pidLine = Get-Content -Path $PidFile -TotalCount 1 -ErrorAction SilentlyContinue
        $pidInt = if ($pidLine) { ("$pidLine").Trim() -as [int] } else { $null }
        if ($pidInt) {
            Stop-Process -Id $pidInt -Force -ErrorAction SilentlyContinue
        }
        Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $LockFile)   { Remove-Item $LockFile -Force -ErrorAction SilentlyContinue }
    Set-Content -Path $StatusFile -Value 'DISCONNECTED' -Encoding ASCII -ErrorAction SilentlyContinue
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
                'Claude for Word RTL - Check for updates',
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
                "Claude for Word RTL Fix is up to date (v$local).",
                'Claude for Word RTL - Check for updates',
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
                'Claude for Word RTL - Check for updates',
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
            $fallbackUrl = 'https://github.com/asaf-aizone/Claude-for-word-RTL-fix/releases/latest'
            [System.Windows.Forms.MessageBox]::Show(
                "Could not check for updates.`n`n" +
                $line.Trim() + "`n`n" +
                "You can check manually at:`n$fallbackUrl",
                'Claude for Word RTL - Check for updates',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to run check-update.js:`n$($_.Exception.Message)",
            'Claude for Word RTL - Check for updates',
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
        "Uninstall Claude for Word RTL Fix?`n`n" +
        "This will:`n" +
        "  - Remove the Startup entry and the Apps and Features registration`n" +
        "  - Stop the tray and the injector`n" +
        "  - Clean node_modules and temp status files`n" +
        "  - Remove any legacy WebView2 env var written by older installs`n`n" +
        "Word itself is not modified. The install folder stays in place -" +
        " you can delete it manually afterward.",
        'Claude for Word RTL - Uninstall',
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $uninstall = Join-Path $InstallDir 'uninstall.bat'
    if (-not (Test-Path $uninstall)) {
        [System.Windows.Forms.MessageBox]::Show(
            "uninstall.bat was not found next to the tray script. Cannot continue.",
            'Claude for Word RTL - Uninstall',
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

# Poll the status file every 2s (matches injector's POLL_MS) and update icon.
# Also detect a stale status: if the injector PID isn't alive AND the status
# file is older than $StaleSeconds, treat the effective status as DISCONNECTED
# regardless of what the file claims. This handles SIGKILL / crash cases where
# the injector couldn't clean up its own status on the way out.
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

    # Connect is available whenever we are NOT currently connected -
    # either Word is down, or it is up but running without the debug flag
    # (Connect will prompt and relaunch). It is also disabled while a
    # Connect flow is mid-flight, to avoid re-entrance.
    #
    # Disconnect is available whenever there is ANY state to tear down:
    # Word running, injector running, or a Connect flow in progress.
    # Previously Disconnect keyed only on Word; if a Connect failed in a
    # way that left the injector running but never launched Word, the
    # user was stranded with Disconnect greyed out and no recovery path.
    $wordRunning = [bool](Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    $connectInProgress = ($script:ConnectState.Phase -ne 'Idle')

    # Auto-launch the injector if Word is up but no injector is attached,
    # and we are not already in the middle of a Connect. Recovery path
    # for the case where the injector crashed while Word (started via
    # Connect) stayed up. If Word was started directly without Connect,
    # it has no debug surface and the relaunched injector just idles.
    #
    # We do NOT check port 9222 here. inject.js polls the port itself
    # every 2s and attaches when WebView2 comes up; launching it early
    # just means it idles harmlessly until the Claude panel opens.
    # Checking port 9222 from the tray every tick would add cost for no
    # gain (Get-NetTCPConnection is a few hundred ms).
    if ($wordRunning -and -not $injectorAlive -and -not $connectInProgress) {
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

    $miConnect.Enabled = (($effective -ne 'CONNECTED') -and (-not $connectInProgress))
    $miDisconnect.Enabled = ($wordRunning -or $injectorAlive -or $connectInProgress)

    if ($effective -eq $script:lastStatus) { return }
    $script:lastStatus = $effective

    if ($effective -eq 'CONNECTED') {
        $tray.Icon = $iconGreen
        $tray.Text = 'Claude for Word RTL - connected'
    } elseif ($effective -like 'ERROR:*') {
        $tray.Icon = $iconRed
        $code = $effective.Substring(6)
        # Human-friendly tooltip mapping. Fall back to the raw code when
        # we have no explicit mapping, so future error codes still render.
        switch ($code) {
            'port-9222-taken-by-other-app' {
                $tray.Text = 'Claude for Word RTL - port 9222 used by another app. Run doctor.bat.'
            }
            default {
                $msg = $code
                if ($msg.Length -gt 55) { $msg = $msg.Substring(0, 55) + '...' }
                $tray.Text = "Claude for Word RTL - error: $msg"
            }
        }
    } else {
        $tray.Icon = $iconRed
        $tray.Text = 'Claude for Word RTL - disconnected'
    }
}
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 2000
$timer.add_Tick($tickAction)
$timer.Start()

# Initial poll so the icon reflects reality immediately instead of after 2s.
& $tickAction

[System.Windows.Forms.Application]::Run()
