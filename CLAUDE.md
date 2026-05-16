# Claude Code instructions for this project

This file is read automatically by Claude Code when a session starts inside this folder. If you are Claude Code helping a user who installed this tool, follow these rules before giving advice.

## What this project is

Client-side RTL fix for Claude's Microsoft Office add-in (Word, Excel, PowerPoint, and Outlook). A small Node.js injector (`scripts/inject.js`) attaches to each Office app's WebView2 host via Chrome DevTools Protocol on a dynamic localhost port and injects CSS + a MutationObserver that fixes Hebrew right-to-left rendering in Claude's task pane. Apache 2.0. Local only, no telemetry. Outlook support is opt-in (v0.3.0+) and gated on a per-launch consent flag - see "Outlook is opt-in" below before touching anything Outlook-related.

**Not an official Anthropic or Microsoft add-in.**

## Codebase map

Top-level files (install folder root):

- `install.bat` - per-user installer. Checks Node, installs npm deps, creates Startup-folder shortcut, writes `HKCU\...\Uninstall\ClaudeWordRTL` (DisplayVersion = "0.3.0" as of this release), silently removes any legacy Auto-enable env var written by v0.1.0 - v0.1.3, stops any old tray/injector via PID files, then launches the new tray. As of v0.1.4 the installer no longer prompts the user for anything. The Apps and Features `DisplayName` is still "Claude for Word RTL Fix" and the Startup-folder shortcut is still `Claude for Word RTL Tray.lnk` for v0.1.x upgrade compatibility, even though v0.3.0 covers four Office apps (Word/Excel/PowerPoint/Outlook).
- `uninstall.bat` - 4-step uninstaller: stop tray and injector, remove Startup entry + Apps and Features key, clear any legacy Auto-enable env var (only when its value matches one of our known strings, `=9222` or `=0`; user-modified values are preserved), prune `node_modules` and temp status files. v0.3.0+ also removes the per-app status JSON (`claude-office-rtl.apps.json`), the Outlook opt-in flag (`claude-office-rtl.outlook-optin`), and the Disconnect-only request file (`claude-office-rtl.disconnect-outlook.request`).
- `doctor.bat` - 19-check diagnostic (was 15 before v0.3.0) for the multi-app architecture. Checks 1-15 are unchanged: Node, npm, deps, per-app Office install detection (Word, Excel, PowerPoint), per-app process-running, dynamic CDP ports (via `node -e` shell-out to `scripts/port-discovery.js`), active Claude target enumeration, injector PID, aggregate status file, per-app status JSON, tray PID, Startup entry, Apps and Features key, the legacy Auto-enable env-var regression check (FAIL if it ever returns), and WebView2 runtime. Checks 16-19 are new and Outlook-specific (all `:info` since Outlook is opt-in): OUTLOOK.EXE installed, Outlook running, Outlook CDP target via dynamic port discovery, Outlook entry in `apps.json`. Writes `doctor.log` next to itself.
- `check-update.bat` - thin wrapper that runs `node scripts/check-update.js` and prints the result. Used by the tray menu and manually from cmd.
- `cleanup.bat` - recovery helper. Kills our current injector via PID file, then scans for orphan `node.exe` processes whose command line contains `inject.js` and stops only those. Also kills tray `powershell.exe` processes whose command line mentions `tray-icon.ps1`. v0.3.0 includes OUTLOOK.EXE in the "is any Office app still running" check so the WebView2-host warning is accurate when Outlook is the only Office app left.
- `start.bat` - older single-step launcher. Pre-tray-era entry point; mostly superseded by the tray UI. Still works for debugging.
- `word-wrapper.bat` - transparent Word launcher invoked by Connect Word and by the document-reopen flow. Sets `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0` in its own process scope (port=0 means "WebView2 picks a free dynamic port", was =9222 in v0.1.x), ensures the hidden injector is running (lock file + PID-alive check to avoid duplicates), then launches `WINWORD.EXE` with an optional document argument.
- `excel-wrapper.bat` - same shape as word-wrapper.bat for `EXCEL.EXE`. Shares `%TEMP%\claude-word-rtl.lock` and `%TEMP%\claude-word-rtl.pid` with the other wrappers; one injector serves all four apps.
- `powerpoint-wrapper.bat` - same shape for `POWERPNT.EXE`. Note: Word and PowerPoint sometimes share a single WebView2 host process, so both apps may register on the same port. The injector handles both shapes via the `_host_Info=` URL parameter.
- `outlook-wrapper.bat` (v0.3.0+) - same shape as the other wrappers for `OUTLOOK.EXE` (classic Outlook only; refuses to run if New Outlook `olk.exe` is detected). Also writes the per-launch opt-in flag (`%TEMP%\claude-office-rtl.outlook-optin`) before launching Outlook; the injector treats Outlook as blocked unless that flag exists. The flag is cleared by the injector at startup so consent does not carry across sessions.
- `inject-hidden.vbs` - runs `node scripts/inject.js` and `powershell.exe ... scripts/tray-icon.ps1` with window style 0 (hidden). Called from any of the four wrappers.
- `install.log` - generated by `install.bat` at install time. Not shipped.
- `CHANGELOG.md`, `README.md`, `README.he.md`, `SECURITY.md`, `LICENSE` - standard project docs.

Files under `scripts/`:

- `scripts/inject.js` - the Node injector. Discovers Office WebView2 hosts at runtime via `port-discovery.js` (no fixed port), attaches to each Claude target it finds, and injects `<style>` + MutationObserver via `Runtime.evaluate`. Re-injects every 2s (`POLL_MS`). Writes PID to `%TEMP%\claude-word-rtl.pid`, aggregate one-line status (`CONNECTED` / `DISCONNECTED` / `ERROR:<msg>`) to `%TEMP%\claude-word-rtl.status`, and per-app state to `%TEMP%\claude-office-rtl.apps.json`. Truncates `%TEMP%\claude-word-rtl.log` on each start.
- `scripts/port-discovery.js` - the v0.2.0 dynamic-port mechanism. Walks `tasklist` for `msedgewebview2.exe` PIDs, maps each PID to its LISTENING port via `netstat`, probes every candidate's `/json/list` for Claude targets, and returns each `{target, app, port}` triple. Per-target app identification reads the `_host_Info=` URL parameter that Office appends to the panel URL.
- `scripts/tray-icon.ps1` - PowerShell tray icon host. Singleton via named mutex. Polls `claude-word-rtl.status` every 2s for the icon color and `claude-office-rtl.apps.json` for the four per-app status labels at the top of the menu (Word/Excel/PowerPoint/Outlook). Implements per-app Connect state machines on Timers (Connect Word / Connect Excel / Connect PowerPoint via a generic loop; Connect Outlook routes through a dedicated `Start-ConnectOutlook` handler with a per-launch warning dialog whose default-focused button is Cancel), Disconnect Outlook only (writes the per-app request file the injector polls), Disconnect-all cleanup (which skips Outlook because it is opt-in - the user may have opened it just to read mail), Show diagnostic log, Check for updates dialog, Uninstall launcher, Exit. Auto-enable toggle was removed in v0.1.4 (security: persistent env var was an EDR trigger). The auto-launch path that brings the injector back up when an Office app is running without one ignores OptIn apps (Outlook) - a bare Outlook session should not trigger CDP attach on its own. Staleness detection: if PID is not alive AND status file is older than `$StaleSeconds`, the effective status is forced to `DISCONNECTED` so the icon does not lie when the injector crashed. Icon glyph is "O" (Office) on a status-colored rounded square with a small RTL arrow.
- `scripts/check-update.js` - fetches `https://api.github.com/repos/asaf-aizone/Claude-for-Office-RTL-fix/releases/latest` via Node's built-in `https` module. Compares `tag_name` to the local `package.json` version numerically. Zero npm dependencies. 5s timeout.
- `scripts/create-shortcut.ps1` - called by `install.bat` to create the Startup-folder `.lnk` that points at `start-tray.vbs`.
- `scripts/start-tray.vbs` - hidden launcher for `tray-icon.ps1`. The Startup-folder shortcut points here.
- `scripts/package.json` - declares `chrome-remote-interface` dependency and the local `version` string. `version` is the authoritative install version for comparisons.
- `scripts/package-lock.json`, `scripts/node_modules/` - generated by `npm install`. Not edited by hand.

Files under `lib/`:

- `lib/office-apps.js` - single source of truth for Office app metadata (name, process executable, `_host_Info=` URL key) PLUS the `BLOCKED_HOST_INFO_KEYS` set of apps that require a per-launch opt-in flag to attach. Outlook is on the block list as of v0.3.0. Used by `scripts/inject.js`, `scripts/port-discovery.js`, and the diagnostic probes. The PowerShell tray maintains its own parallel table (`$Apps` in `tray-icon.ps1`, with an `OptIn = $true` marker on the Outlook entry) since it cannot `require` a Node module; the two are kept in sync by hand.

Other directories:

- `docs/` - supplementary docs. Currently contains `docs/security.md` (full threat model) and `docs/images/` (README screenshots: tray icon states, tray menu, installer output, before/after, Connect dialog).
- `probe/` - POC scripts used during the v0.2.0 architecture work and kept in-tree for diagnosing real-machine issues. `probe/launch-office-dynamic.bat` and `probe/dynamic-port-discovery.js` validate dynamic-port discovery; `probe/text-rendering-tests.js` is a 15-variation RTL+typography suite per app, run via CDP.
- `docs/bugs/` - does not exist in the current tree. If you encounter references elsewhere, it is a historical or future directory.
- `test-results/` - does not exist in the current tree. This project has no automated test harness checked in.

## Architecture in one page

End-to-end execution flow when a user right-clicks the tray and picks Connect Word (Excel and PowerPoint follow the identical flow):

1. User right-clicks the tray icon near the clock. `scripts/tray-icon.ps1` owns the icon and menu. The menu shows three disabled status labels at the top (Word/Excel/PowerPoint, each as "connected", "not running", "running without RTL", or "error"), three Connect items below them (Connect Word / Connect Excel / Connect PowerPoint), and a single Disconnect-all that closes every Office app and the injector.
2. Click-handler on Connect Word starts the per-app Connect state machine on a `System.Windows.Forms.Timer`. The menu closes immediately so the UI thread is never blocked. Each tick advances a phase: enumerate open Word documents via the `Word.Application` COM ProgId and the `Documents` collection, confirm with the user, gracefully close Word, wait up to 10s, then launch the wrapper. Excel uses `Excel.Application` + `Workbooks`, PowerPoint uses `PowerPoint.Application` + `Presentations`. Same state machine, different per-app metadata.
3. If the Office app does not close within 10s, a force-close OK/Cancel dialog is shown. OK kills the process by name, Cancel aborts the Connect.
4. The wrapper call is `word-wrapper.bat "<doc path>"` (or no arg). The batch file sets `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0` in its own env scope. Port=0 means WebView2 picks a free dynamic port per process, so multiple Office apps launched through their respective wrappers each get their own debug surface without colliding. Excel and PowerPoint wrappers behave identically, with their own EXE locator and document/workbook/presentation argument.
5. The wrapper checks `%TEMP%\claude-word-rtl.lock` + `%TEMP%\claude-word-rtl.pid`. If that PID is alive, it skips launching the injector (idempotent across all three wrappers). Otherwise it writes the lock file and runs `wscript inject-hidden.vbs`.
6. `inject-hidden.vbs` runs `cmd /c node scripts\inject.js` with window style 0 and also re-launches `scripts\tray-icon.ps1` hidden (the tray's singleton mutex will silently exit the duplicate).
7. The wrapper then `start`s the Office EXE with the document argument. The Office app starts, inherits the env var, WebView2 reads it and opens a CDP port on a dynamic localhost port number.
8. `scripts/inject.js` enumerates Office WebView2 ports each tick via `port-discovery.js`: `tasklist` for `msedgewebview2.exe` PIDs, `netstat` for the LISTENING TCP ports those PIDs own, probe each port's `/json/list`, and identify each Claude target's app via the `_host_Info=` URL parameter (Word, Excel, or Powerpoint). When it finds a Claude target it has not seen before, it attaches via WebSocket, calls `Runtime.evaluate` with `INJECTOR_SCRIPT`, and writes `CONNECTED` to the aggregate status file plus the per-app entry to `claude-office-rtl.apps.json`. The MutationObserver inside the page keeps RTL CSS applied across Claude's client-side re-renders and performs typography cleanup (em-dash and en-dash become hyphen, arrow glyphs become comma) on text nodes only, skipping `<pre>`, `<code>`, `<kbd>`, `<samp>`, textareas, and `contenteditable`.
9. 5s after each inject, the injector runs a DOM validation pass. If `direction: rtl` did not apply or known Claude-panel selectors no longer match, it writes `ERROR:dom-not-matched-after-inject` and logs a `[WARN]`. The tray turns red on that ERROR.
10. The tray's 2s tick reads the aggregate status file, cross-checks the injector PID is alive, parses `claude-office-rtl.apps.json` to update the three per-app status labels, and updates the icon color. If PID is dead and the status file is stale (`$StaleSeconds`), the effective status is forced to `DISCONNECTED` regardless of what the file claims - guard against SIGKILLed injectors leaving stale `CONNECTED` behind.

Auto-enable removal in v0.1.4 (preserved unchanged in v0.2.0):

- v0.1.0 - v0.1.3 had an "Auto-enable" tray checkbox that wrote
  `HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222`
  so every Word launch had RTL ready without clicking Connect. That variable is
  read by every WebView2 host on the account (Teams, Outlook, Edge WebView,
  OneDrive UI), and enterprise EDR products (Microsoft Defender for Endpoint,
  CrowdStrike, SentinelOne) flagged the modification as a credential-theft
  signal. After a field incident with EDR-driven host isolation, v0.1.4
  removed Auto-enable entirely. The tray no longer has the checkbox, the
  installer no longer prompts, and the env var is never written. v0.2.0 does
  not bring it back. Activation is Connect-only across all three Office apps:
  user clicks Connect Word/Excel/PowerPoint from the tray, the matching
  wrapper sets the WebView2 debug flag in its own process scope, the Office
  app inherits it.
- The tray still auto-launches the injector when it notices any Office app is
  running without a live injector. 30s cooldown so a missing-Node or
  crash-on-startup condition does not spin forever. Covers the recovery path
  (injector died while an Office app stayed up after Connect).
- A reinstall over v0.1.x silently removes the legacy env var if its value
  matches one of our known strings (`=9222` or `=0`). User-modified values
  are preserved.

Status file contracts:

- `%TEMP%\claude-word-rtl.status` - aggregate state. Single line, ASCII. One of `CONNECTED`, `DISCONNECTED`, or `ERROR:<message>`. Drives the tray icon color. Written by `scripts/inject.js` on every state change. Read by `scripts/tray-icon.ps1` every 2s. Truncated (not appended).
- `%TEMP%\claude-office-rtl.apps.json` - per-app state (v0.2.0+). Single JSON object: `{"Word":"CONNECTED","Excel":"DISCONNECTED","PowerPoint":"CONNECTED"}`. Drives the three per-app status labels at the top of the tray menu. Atomic write (temp + rename) so the tray never reads a half-written object. Written by `scripts/inject.js` after every attach/disconnect. Missing file means "all DISCONNECTED" to the tray.

Update flow (`Check for updates...`):

- Tray invokes `node scripts/check-update.js`.
- `check-update.js` GETs `https://api.github.com/repos/asaf-aizone/Claude-for-Office-RTL-fix/releases/latest` with a 5s timeout. User-Agent is set, `Accept: application/vnd.github+json`.
- Script parses the JSON, normalizes `tag_name`, compares numerically to `require('./package.json').version`. Pre-release suffixes (`-beta` etc.) are stripped.
- If newer: prints `UPDATE_AVAILABLE` + version + download URL. The tray dialog shows the current install folder path, opens the GitHub release page in the default browser, and opens the install folder in Explorer.
- If up to date or network error: tray shows a simple dialog. Manual fallback link is included on error.

## Common commands

Install:

```
install.bat
```

Run the installer by double-clicking `install.bat` (or from cmd in the install folder). No admin rights. Runs 4 steps with no prompts. As of v0.1.4, the installer also silently removes any legacy Auto-enable env var written by v0.1.0 - v0.1.3.

Uninstall:

```
uninstall.bat
```

Or right-click the tray > `Uninstall...`. Both run the same script.

Diagnostics:

```
doctor.bat
```

Runs 19 checks (was 15 before v0.3.0) and writes `doctor.log`. Attach this log when reporting an issue. Notable v0.2.0 checks: dynamic CDP ports discovered (#6), active Claude targets per app (#7), per-app injector status from `apps.json` (#10), and the critical legacy-env-var regression check (#14, FAIL if `HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` ever returns). v0.3.0 adds Outlook-specific checks (#16 install, #17 process, #18 CDP target, #19 apps.json entry) - all `:info` because Outlook is opt-in.

Check for updates:

```
check-update.bat
```

Or right-click the tray > `Check for updates...`. The tray version shows a richer dialog with the install folder path.

Manually start the tray (if it crashed or the Startup entry was removed):

```
wscript scripts\start-tray.vbs
```

Runs hidden. Singleton mutex ensures only one tray at a time.

Force cleanup (nuclear option for stuck state):

```
cleanup.bat
```

Kills the injector via PID file, scans for orphan `node.exe` running `inject.js`, kills stray `tray-icon.ps1` PowerShell processes, and includes WINWORD/EXCEL/POWERPNT/OUTLOOK in the "is any Office app still up" check (Outlook coverage added in v0.3.0).

Smoke test after a code change (manual, since there is no automated suite):

1. `uninstall.bat`, then `del %TEMP%\claude-word-rtl.* %TEMP%\claude-office-rtl.*`.
2. `install.bat` fresh. Confirm `install.log` has no WARN/FAIL.
3. `doctor.bat` and confirm `doctor.log` is all `[OK]` and expected `[INFO]`. Step 16-19 are `[INFO]` even when Outlook is not installed.
4. Open Word with a Hebrew document; right-click tray, Connect Word; confirm RTL is applied.
5. Repeat for Excel and PowerPoint - all three apps should connect at the same time, each with its own per-app status label going CONNECTED in the tray menu.
6. (v0.3.0+) Right-click tray, Connect Outlook. Confirm the warning dialog appears with Cancel as the default-focused button. OK to proceed; Outlook relaunches via `outlook-wrapper.bat`, opt-in flag appears at `%TEMP%\claude-office-rtl.outlook-optin`, Outlook status flips to CONNECTED. Pick Disconnect Outlook only and confirm only Outlook drops (Word/Excel/PowerPoint stay green).
7. Right-click tray, Disconnect all. Confirm icon goes red, Word/Excel/PowerPoint close, Outlook stays open (Disconnect-all skips OptIn apps). Outlook opt-in flag is gone.
8. `check-update.bat` reports the current version as latest.

For the full M5 release smoke-test scenarios (Connect during dialog, force-close paths, EDR verification, legacy-user upgrade behaviour, etc.), see section 8 of `docs/OUTLOOK-EXPANSION-PLAN.md`.

Log locations:

- `%TEMP%\claude-word-rtl.log` - injector diagnostic log, truncated on each injector start.
- `install.log` - written next to `install.bat` at install time.
- `doctor.log` - written next to `doctor.bat` when it runs.

## State locations

Every file and registry key the tool creates or reads:

- `%TEMP%\claude-word-rtl.status` - aggregate one-line status (`CONNECTED` / `DISCONNECTED` / `ERROR:<msg>`). Written by `scripts/inject.js`, read by `scripts/tray-icon.ps1`. Drives icon color.
- `%TEMP%\claude-office-rtl.apps.json` - per-app state (v0.2.0+) for the four status labels in the tray menu (Word/Excel/PowerPoint/Outlook). Written atomically by `scripts/inject.js` after every attach/disconnect.
- `%TEMP%\claude-office-rtl.outlook-optin` (v0.3.0+) - per-launch opt-in flag for Outlook CDP attach. Existence = the user has consented for the current injector session. Written ONLY by `outlook-wrapper.bat` after the user clicks Connect Outlook and OKs the warning dialog. Cleared by `inject.js` at startup, by the 15-min auto-disconnect timer, by Disconnect Outlook only, by Disconnect all, and by uninstall. Never persisted across reboots.
- `%TEMP%\claude-office-rtl.disconnect-outlook.request` (v0.3.0+) - IPC request file from the tray to the injector for "Disconnect Outlook only". The tray writes a zero-byte file; the injector polls each tick, closes Outlook CDP clients, revokes the opt-in flag, and deletes the request file. Cleared at injector startup so a stale request from a prior session does not fire against a fresh attach.
- `%TEMP%\claude-word-rtl.pid` - injector process ID, one line. Written on injector start, removed on graceful exit. Used by cleanup, uninstall, the tray's staleness check, and all three wrappers' anti-duplicate logic.
- `%TEMP%\claude-word-rtl.tray.pid` - tray process ID, one line. Written on tray start, removed on graceful exit or uninstall. Used by `install.bat` to stop the previous tray during reinstall so the new `tray-icon.ps1` code actually loads.
- `%TEMP%\claude-word-rtl.lock` - written by any of the three wrappers when one of them launches the injector. Prevents double-spawn when the user opens Excel right after Word, etc. Cleaned by `cleanup.bat` or by the wrapper itself on staleness. Filename keeps the `claude-word-rtl` prefix for v0.1.x upgrade compatibility.
- `%TEMP%\claude-word-rtl.log` - rolling diagnostic log from the injector. Truncated at each injector start.
- `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Claude for Word RTL Tray.lnk` - Startup-folder shortcut pointing at `scripts\start-tray.vbs`. Created by `install.bat`, removed by `uninstall.bat`. Filename kept for v0.1.x upgrade compat. User can delete manually via `shell:startup`.
- `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL` - Apps and Features registration so the tool appears in Windows Settings > Apps > Installed apps. Values: `DisplayName` (still "Claude for Word RTL Fix" for v0.1.x upgrade compat), `DisplayVersion`, `Publisher`, `InstallLocation`, `UninstallString`, `DisplayIcon`, `URLInfoAbout`, `NoModify`, `NoRepair`. Written on install, removed on uninstall.
- `HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` - **NOT written by v0.1.4 or v0.2.0**. v0.1.0 - v0.1.3 wrote this when Auto-enable was on; v0.1.4+ install.bat and uninstall.bat both clear it (only when its value matches a known legacy string, `=9222` or `=0`; user-modified values are preserved). v0.2.0 sets the WebView2 debug flag per-process via the matching wrapper instead, so other WebView2 hosts on the account are never affected.
- CDP target the injector looks for: `pivot.claude.ai` (primary URL pattern) or any `*.claude.ai` (fallback). See `scripts/inject.js`.

## Platform - Windows only

This tool is Windows-only by design. It will not run on macOS or Linux.

- The core mechanism is CDP attach to Microsoft WebView2 on a localhost dynamic port. WebView2 is Windows-only. Office for Mac uses WKWebView instead, which does not expose the same interface.
- The launcher stack (install.bat, uninstall.bat, start.bat, word-wrapper.bat, excel-wrapper.bat, powerpoint-wrapper.bat, doctor.bat, check-update.bat, cleanup.bat, inject-hidden.vbs, start-tray.vbs, create-shortcut.ps1, tray-icon.ps1) is native Windows (Batch, VBScript, PowerShell with System.Windows.Forms).
- HKCU registry, the Startup folder, `%TEMP%` and `%APPDATA%` are Windows concepts with no cross-platform equivalent in this codebase.

If a user asks whether this works on Mac, say no and explain the WebView2 vs WKWebView gap. Do not suggest porting. Do not suggest Wine or a Windows VM to run the tool on Mac hardware - the tool runs fine there mechanically, but they still need Office for Windows and the Claude add-in to be the Windows variant, and at that point they are effectively a Windows user. Point them at Office for Windows if they need RTL in Claude's add-in.

## Why this file exists

This project ships a `CLAUDE.md` because troubleshooting it well requires reading files the AI assistant cannot see from a remote chat: `%TEMP%\claude-word-rtl.log`, `%TEMP%\claude-word-rtl.status`, `%TEMP%\claude-office-rtl.apps.json`, `doctor.log`, and the actual state of the install folder. Claude Code (the local CLI/IDE agent) reads those files and can run `netstat`, `tasklist`, and `curl`; Claude Chat on the web cannot. When recommending this tool to users, also recommend that for support they open a Claude Code session inside the install folder, not a Chat window. The difference in diagnostic accuracy is large.

## Before answering upgrade or troubleshooting questions, do this first

The user may be on an older version with known bugs. Always check current state before advising.

1. **Read the installed version:** `scripts/package.json` field `"version"`.
2. **Fetch the latest release notes from GitHub** - do NOT rely only on this file, which ages:
   - Releases: <https://github.com/asaf-aizone/Claude-for-Office-RTL-fix/releases>
   - Latest release API (JSON): <https://api.github.com/repos/asaf-aizone/Claude-for-Office-RTL-fix/releases/latest>
   - Use `gh release view --json name,tagName,body` if the user has `gh` installed, otherwise `curl` or `WebFetch` to the API URL.
3. If the installed version is older than the latest, read the release notes for every version in between. Release notes describe bugs, fixes, and upgrade steps.
4. Read `CHANGELOG.md` in this folder for the complete history.
5. Read the README - specifically the "Updating to a newer version" / "עדכון לגרסה חדשה" section - before walking the user through an upgrade.

Note: the GitHub repository was renamed in v0.2.1 from `Claude-for-word-RTL-fix` to `Claude-for-Office-RTL-fix` to match the v0.2.0 multi-app product name. GitHub serves a permanent HTTP 301 redirect from the old URL, so existing clones, bookmarks, v0.1.x install links, and `git remote` URLs all keep working without intervention. v0.2.1 also taught `scripts/check-update.js` to follow 301/302 redirects so future repository moves cannot break the update notification path.

## Known issue in v0.1.0 (fixed in v0.1.1)

If the user says the tray is **red and stays red even after clicking "Connect"** or "nothing happens when I open Word", check `scripts/package.json`:

- If version is `0.1.0`: the user has a known bug. Auto-enable sets the WebView2 debug flag on Word but does NOT start the Node injector, so the tray has nothing to attach to. Tell the user to upgrade to v0.1.1 or newer - the fix is that the tray now auto-launches the injector whenever Word is running without one. The v0.1.0 "Check for updates" dialog does not show the install folder; users have to find it manually via **Windows Settings > Apps > Installed apps > Claude for Word RTL Fix**. Full upgrade steps in the v0.1.1 release notes.
- If version is `0.1.1` or later (current is `0.3.0`): the auto-launch mechanism is in place. Walk through tray log diagnostics: `%TEMP%\claude-word-rtl.log`, `%TEMP%\claude-word-rtl.status`, `%TEMP%\claude-office-rtl.apps.json`, `%TEMP%\claude-word-rtl.pid`. If the report involves Outlook, also check `%TEMP%\claude-office-rtl.outlook-optin` (existence = the user has clicked Connect Outlook this session) and look for the blocked-target line in the log when the flag is absent.

## What NOT to do

- Do not modify Office.js, Word/Excel/PowerPoint templates, or any Office add-in registration. This tool does not touch them.
- Do not run `uninstall.bat` to "fix" a broken install unless the user explicitly asks to uninstall - it clears the legacy Auto-enable env var, Apps and Features registration, and `node_modules`. Reinstalling over the top via `install.bat` is almost always the right recovery path.
- Do not recommend setting the debug port flag on Edge/Chrome/Teams directly - the flag is for Office's WebView2 hosts only, set per-process by the three wrappers.
- Do not suggest the user disable Office add-in security, SmartScreen, or Defender. The tool does not require that.
- Do not help the user force-push to `main` or rewrite history. Contributions come via PRs.
- Do not suggest porting this tool to macOS or Linux, or recommend running it under WSL, Wine, or a VM as a workaround. The CDP port does not exist on Office for Mac because Office for Mac does not use WebView2. See "Platform - Windows only" above.
- Do not auto-connect Outlook for the user, write the opt-in flag on their behalf, or remove the warning dialog "to make Connect Outlook smoother". The opt-in model is intentional - mail content enters the panel DOM during Summarize/Draft, and that is qualitatively more sensitive than Word/Excel/PowerPoint document content. The flag, the dialog, and the 15-min auto-disconnect are all gates the user must cross consciously. If a user complains "Connect Outlook is annoying", the answer is "use Disconnect Outlook only when you finish a mail session, not the dialog removal".
- Do not modify Outlook accounts, signatures, mail rules, or PST/OST files. This tool does not touch any of them.
- Do not suggest enabling New Outlook (`olk.exe`) support. M0 deferred it; the wrapper and Connect Outlook flow both refuse if `olk.exe` is detected. If a user wants RTL in New Outlook, point them at the deferred status in `docs/OUTLOOK-EXPANSION-PLAN.md` section 7 and the M0 findings in `probe/README.md`.

## When the user asks "what version is latest?"

Do not answer from memory. Fetch <https://api.github.com/repos/asaf-aizone/Claude-for-Office-RTL-fix/releases/latest> and read the `tag_name` and `body`. Models trained before the latest release will not know.
