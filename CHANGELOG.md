# Changelog

All notable changes to this project will be documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [0.1.4] - 2026-05-01

Security hotfix. No new features. If you are using v0.1.0 -
v0.1.3, upgrading to v0.1.4 is strongly recommended,
particularly on managed corporate devices.

### Removed

- Persistent Auto-enable user environment variable.
  Earlier versions wrote
  `HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`
  with the WebView2 debug flag so every Word launch on the
  account picked up the flag automatically. The same
  variable is read by every other WebView2 host on the
  account (Microsoft Teams, the Outlook reading pane, Edge
  WebView, the OneDrive UI, etc.), which made the change
  far broader than its name suggested. Enterprise EDR
  products (Microsoft Defender for Endpoint, CrowdStrike
  Falcon, SentinelOne, Sophos) treat unexpected
  modifications of WebView2 browser arguments as a
  credential-theft signal and may trigger host isolation on
  managed devices when this pattern fires alongside the
  other patterns the installer produces (hidden VBS
  launcher, PowerShell with `-ExecutionPolicy Bypass`,
  autoruns from a non-trusted install path). v0.1.4 removes
  the toggle entirely. The WebView2 debug flag is now set
  in `word-wrapper.bat`'s process scope only, inherited by
  Word but not by Teams/Outlook/Edge.
- Tray menu's "Auto-enable at every Word launch" checkbox.
  All RTL activation now flows through Connect, which uses
  `word-wrapper.bat` to set the flag per-process.

### Changed

- `install.bat` no longer prompts for Auto-enable at the
  end of installation. Step count goes from 5 to 4. On a
  reinstall over v0.1.0 - v0.1.3, it silently removes the
  legacy `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222`
  registry value, so users coming from older versions are
  migrated automatically. A user-modified value is preserved.
- README.md and README.he.md now lead with a prominent
  warning for managed corporate devices, advising users
  not to install without prior approval from their security
  team.

### Security

- This release exists because of a field incident. A
  v0.1.3 install on a corporate-managed Windows machine
  was flagged by EDR as a token-theft attempt and the
  device was placed in network isolation. Root cause was
  the user-scope env var setting; once removed, the
  remaining patterns alone are still detectable but no
  longer trigger automatic isolation in the products we
  observed. Managed devices may still surface SmartScreen
  or generic-untrusted-software warnings; users should
  obtain explicit IT approval before installing.

### Upgrade notes

Run `install.bat` over your existing install. The legacy
env var is removed automatically. After upgrade, you must
use the tray Connect menu once per Word session for RTL,
even if you previously had Auto-enable on. There is no
longer a way to make Word launch with RTL automatically
from the taskbar or from a `.docx` double-click; this is
intentional, see Removed section.

## [0.1.3] - 2026-04-23

### Fixed

- Tray stays red on a clean install with Auto-enable on, even when
  Word is running and Node is running. Root cause: WebView2 opens
  its CDP debug server on `[::1]:9222` (IPv6 loopback) while other
  apps - notably Google Drive File Stream - bind to
  `127.0.0.1:9222` (IPv4). Node's `http.get('http://localhost:...')`
  resolves `localhost` to IPv4 first and hits whoever got there
  first, so `listTargets()` was talking to Google Drive and seeing
  no Claude panel. `scripts/inject.js` now probes both
  `127.0.0.1:9222` and `[::1]:9222` explicitly and merges the
  target lists, so the injector finds the Word panel regardless of
  which family Windows resolves `localhost` to and regardless of
  what else is listening on the other family.
- `scripts/inject.js` now distinguishes "no targets on either
  family" (normal during Word startup) from "targets exist on
  9222 but none of them are claude.ai" (another app owns the
  port). In the second case it writes
  `ERROR:port-9222-taken-by-other-app` to the status file so the
  tray tooltip names the actual problem instead of showing a
  generic disconnected state. Users chasing a stuck-red tray get
  pointed at the port conflict immediately rather than reinstalling
  in circles.

### Added

- `doctor.bat` now runs explicit IPv4 and IPv6 probes against port
  9222 (`127.0.0.1:9222` and `[::1]:9222`) and a
  `netstat -ano | findstr :9222` dump that surfaces which PID owns
  which address family. When the injector cannot attach, the log
  immediately shows whether the CDP port is split across families
  or owned by a non-Word process, which is the fastest path from
  "tray is red" to "Google Drive is squatting on IPv4 port 9222".
- Troubleshooting row in both README.md (English section) and
  README.he.md for "tray stays red despite Auto-enable on and
  Node.js running". Points the user at `doctor.bat` and the
  `netstat` check, and names Google Drive File Stream as a known
  offender that grabs `127.0.0.1:9222` on some machines. Saves
  users the diagnostic trip when the symptom matches.
- Windows-only notices across the project docs, added during this
  session between the 0.1.2 tag and 0.1.3: a Hebrew blockquote and
  an English blockquote at the top of the README stating the tool
  runs on Windows only, strengthened "Mac not supported" bullets
  in the prerequisites list, a Windows-only callout in
  README.he.md, a new "Platform - Windows only" section in
  `CLAUDE.md` explaining the WebView2 vs WKWebView gap, and a
  "What NOT to do" bullet telling Claude Code sessions not to
  suggest Wine, a Windows VM, or a macOS/Linux port as a
  workaround. Cuts off a recurring support thread where Mac users
  asked whether a tweak or a VM could make it work.
- Troubleshooting sections in README.md (Hebrew and English) and
  README.he.md now lead with a "use Claude Code, not Claude Chat"
  callout. Claude Code runs locally and reads
  `%TEMP%\claude-word-rtl.log`, `doctor.log`, and project
  `CLAUDE.md` directly, and can execute `netstat` and `curl` to
  identify port-owner conflicts. Claude Chat cannot see those
  files and ends up guessing blind. A recent support case dragged
  on because the user debugged via Chat and only got generic
  steps; Code would have pinpointed the Google Drive File Stream
  port-9222 conflict from the log in the first turn.

### Changed

- `scripts/inject.js` `listTargets()` no longer assumes a single
  loopback address. The function now issues parallel requests to
  the IPv4 and IPv6 loopback endpoints, tolerates one side failing
  (the common case - only one family is actually listening), and
  de-duplicates targets by `id` before returning. Behavior is
  unchanged on machines where only one family is in use; the fix
  only matters when two apps split the port across families.

## [0.1.2] - 2026-04-23

### Added

- `CLAUDE.md` at the repo root. Claude Code sessions launched inside
  the install folder read it automatically and get pointers to the
  release notes, the "red tray" known issue for v0.1.0, the tray/
  injector file layout under `%TEMP%`, and rules about what NOT to
  recommend. Purpose: when a user asks Claude Code on their machine
  for help with this tool, the agent knows to fetch the latest
  release notes before advising, instead of reasoning from stale
  training data.

## [0.1.1] - 2026-04-23

### Fixed

- Tray stays red indefinitely when Auto-enable is on. Auto-enable sets
  the WebView2 debug flag on Word but does not start the Node injector;
  `word-wrapper.bat` is the only thing that did, and the wrapper only
  runs via the tray's Connect action. A user who followed the
  "Auto-enable = set-and-forget" recommendation would open Word at
  login, see the tray red, click Connect (which closes and reopens
  Word), and wonder why Auto-enable did not save them the Connect step.
  The tray now auto-launches the injector when it sees Word running
  without a live injector, respecting a 30s cooldown to avoid a spin if
  Node is missing or crashing on startup. Covers both the Auto-enable
  path (no wrapper) and the recovery path (injector died while Word
  stayed up).
- `install.bat` :log helper truncated or broke on any logged string
  containing a closing parenthesis ")". The if-else body used cmd
  parens, and cmd treats ")" inside a parens block as the block
  terminator regardless of quoting. Rewritten without parens, using
  labels, so logging arbitrary human text is safe.
- Reinstall-over-top upgrade path did not actually replace the running
  tray. The tray's singleton mutex caused the freshly-launched tray to
  exit silently, leaving the OLD tray-icon.ps1 code in memory until
  logout. Users who applied a patch via "Check for updates" kept
  seeing the pre-patch behavior. `install.bat` now stops the previous
  tray (and injector) via their PID files before starting the new
  tray, so the upgrade takes effect immediately.

### Changed

- Tray "Check for updates..." dialog, when an update exists, now shows
  the install folder path and opens Explorer there in addition to the
  download page. Previously the dialog pointed at the browser download
  only, leaving the user to remember where they installed and extract
  in the right place - a common source of failed upgrades.
- README has a new "Updating to a newer version" subsection (Hebrew
  and English) spelling out the three-step upgrade: extract over,
  close Word, run install.bat.

### Changed

- `install.bat` now has a fifth step that prompts the user to turn on
  Auto-enable at install time (recommended, Y/N prompt). On reinstall,
  the step detects that the variable is already set to our exact value
  and skips the prompt silently. On a conflicting value, the warning is
  folded into the same dialog so the user can make an informed choice.
  The purpose is to get users out of the clunky close-and-reopen-Word
  Connect flow for every session - with Auto-enable on, Word just opens
  with RTL ready.
- `install.bat` Node.js prerequisite check is louder on failure: a full
  how-to (download LTS from nodejs.org, install with defaults, verify
  with `node --version`, re-run install.bat) instead of a terse
  one-liner. Also warns when Node.js is present but below the
  recommended v16 floor.
- README install sections (Hebrew + English) lead with a dedicated
  Node.js callout before the numbered steps, and the steps now begin
  with "install Node.js" explicitly. The Auto-enable prompt at the end
  of install.bat is mentioned in the step list so users expect it.

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
