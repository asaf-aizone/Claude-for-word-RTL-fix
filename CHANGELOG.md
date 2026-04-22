# Changelog

All notable changes to this project will be documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [0.1.0] - 2026-04-21

Initial public release.

### Added

#### RTL fix
- CSS injection via Chrome DevTools Protocol on `localhost:9222` that
  sets `direction: rtl` on Claude's Word add-in panel, fixes list
  markers to render on the right, keeps code blocks LTR, and isolates
  inline code inside Hebrew paragraphs so bidi flow stays correct.
- MutationObserver in the panel that replaces em-dash and en-dash
  with hyphen, and arrows (`\u2190-\u2199`, `\u21D0-\u21D4`, HTML
  entities) with comma, in visible text only. Skips `<pre>`,
  `<code>`, `<kbd>`, `<samp>`, text areas, and `contenteditable`
  nodes so source code and user input are never modified.
- DOM validation 5 seconds after each injection: verifies
  `direction: rtl` actually applied and known Claude-panel selectors
  still match. On failure, emits a `[WARN]` and writes
  `ERROR:dom-not-matched-after-inject` to the status file so the
  tray turns red - early warning when Claude changes its DOM.

#### Install model (tray-only)
- `install.bat` - per-user installer (no admin required). Creates a
  Startup-folder entry pointing at `scripts\start-tray.vbs` and
  registers the tool in Windows Settings > Apps > Installed apps
  (HKCU uninstall key, `DisplayName` "Claude for Word RTL Fix").
  No HKCU file associations, no Start Menu shortcut - the tray is
  the single entry point.
- `uninstall.bat` - removes the Startup entry, stops the tray and
  the injector via their own PID files, removes the Apps and
  Features registration, removes the Auto-enable environment
  variable if it matches our exact value, cleans `node_modules`
  and temp status files.
- `doctor.bat` - 12-check diagnostic script covering Node, npm,
  dependencies, Word install, Word running, debug port 9222,
  injector (via PID file), Startup entry, tray process (via PID
  file), WebView2, Office version, and the Apps and Features
  registration. Writes `doctor.log`.
- `check-update.bat` + `scripts/check-update.js` - checks the
  GitHub releases API for a newer version, numeric component
  comparison. Zero npm dependencies - uses Node's built-in
  `https` module only.

#### Tray UX
- `scripts/tray-icon.ps1` - PowerShell tray icon that reflects
  injector status: green (connected), red (disconnected or error),
  gray (starting). Singleton enforced via global mutex. Writes
  `%TEMP%\claude-word-rtl.tray.pid` on start, clears on exit so
  uninstall and doctor target only our process.
- Tray menu:
  - Connect (relaunch Claude for Word RTL Fix) - non-blocking
    `Timer`-driven state machine. Enumerates open documents via COM
    before closing Word, confirms, gracefully closes, waits up to
    10s, then relaunches via the wrapper with the same documents
    reopened. Never stalls the UI thread; the menu dismisses
    normally.
  - Disconnect (close Claude for Word RTL Fix) - closes all Word
    windows gracefully, force-kills any that refused to close.
  - Auto-enable at every Word launch (toggle) - when on, sets a
    per-user HKCU environment variable
    `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222`
    and broadcasts `WM_SETTINGCHANGE` so new processes see it
    immediately. Every new Word launch (taskbar, Recent, docx
    double-click, email attachment) has the RTL fix available
    without requiring a Connect. Confirmation dialog in both
    directions; conflict with an existing non-matching value is
    surfaced in the same dialog. Uncheck or uninstall removes the
    variable, but only when it matches our exact value - never
    clobbers a variable the user set for another purpose.
    Heads-up in the dialog: the variable is user-scoped and may be
    read by other WebView2 apps on the account.
  - Show diagnostic log - opens `%TEMP%\claude-word-rtl.log` in
    notepad.
  - Check for updates... - runs `scripts/check-update.js` and shows
    the result in a dialog. On a newer release, a one-click button
    opens the download page in the default browser; on network
    error or a pre-release 404, a warning dialog surfaces the raw
    message plus a manual fallback link. The menu item is disabled
    while the check runs to prevent double-runs.
  - Uninstall... - runs `uninstall.bat` directly from the tray.
    The tray releases its mutex and cleans up its own state before
    detaching the uninstall process.
  - Exit - shuts down the tray without uninstalling.
- Force-close fallback in Connect: when Word does not exit
  gracefully within 10 seconds, an OK/Cancel dialog offers to
  force-kill `winword.exe` (with an explicit unsaved-changes
  warning) or leave Word alone.

#### Word launch wrapper
- `word-wrapper.bat` - transparent launcher invoked by Connect
  (and the tray's document-reopen flow). Sets the WebView2 debug
  flag, launches the hidden injector via VBS if not already running
  (lock file + PID-alive check avoid duplicates), then launches
  Word with an optional document argument.
- `inject-hidden.vbs` - runs the Node injector without a visible
  window. Also launches the tray PowerShell host in hidden mode.
- `scripts/start-tray.vbs` - hidden launcher for the tray;
  pointed at by the Startup entry.
- `scripts/inject.js` - CDP injector. Finds the Claude panel
  target on `localhost:9222`, injects RTL CSS + MutationObserver,
  re-attaches on panel reload. Writes its PID to
  `%TEMP%\claude-word-rtl.pid`. Writes rolling diagnostic log to
  `%TEMP%\claude-word-rtl.log` on every state change (targets
  discovered, attach errors with stack traces, `listTargets`
  failures). Log is truncated at the start of every run.

### Known limitations

- WebView2's debug flag is read only when a Word process starts.
  A Word session launched without the flag cannot be attached to
  mid-flight; Connect closes and relaunches Word. Enable
  Auto-enable to make this automatic for all future launches.
- Depends on the current DOM structure of Claude's Word add-in
  (still in beta). A Claude update can break CSS selectors or URL
  matching without warning. The injector emits a DOM-validation
  warning in that case; expect occasional releases to track
  upstream changes.
- Does not affect text written directly into the Word document by
  Claude via Office.js (a separate code path not reachable from
  CDP-injected page scripts).
- Tested on Windows 10/11 with Microsoft 365 (Office 16). Volume
  licensed / LTSC builds may behave differently and are not
  currently tested.
