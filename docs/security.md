# Security and Threat Model

This document explains what the tool does at runtime, what risks exist, and what it specifically does not do. Read this before using in a sensitive environment.

## Entry points

The tool can be started in two ways, both of which end up with the same pair of running processes (Word + a hidden Node injector):

1. **Tray icon (recommended).** After `install.bat` runs, a per-user Startup-folder shortcut (`Claude for Word RTL Tray.lnk`) launches `scripts/start-tray.vbs` at login, which launches the tray icon. The user opens Word normally and then right-clicks the tray icon and picks **Connect**; the tray enumerates the open documents via COM, gracefully closes Word, and relaunches it through `word-wrapper.bat` with the debug flag set and the same documents reopened. No registry writes, no file-association changes. Fully reversed by `uninstall.bat`.
2. **`start.bat`** - debug mode. Launches Word and the injector in the foreground with a visible log window. Useful for troubleshooting. Leaves no persistent state behind.

## What happens at runtime

1. `word-wrapper.bat` (or `start.bat`) sets the environment variable `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222`. This is a Microsoft-documented WebView2 flag used for add-in debugging.
2. The wrapper launches Word. Word's WebView2 inherits the flag and exposes Chrome DevTools Protocol on `localhost:9222`.
3. `scripts/inject.js` runs (hidden via `inject-hidden.vbs`, or visibly via `start.bat`). It opens an HTTP connection to `http://localhost:9222/json/list`, finds the page target whose URL matches the Claude panel (`pivot.claude.ai`, fallback `claude.ai`), and opens a WebSocket to that target's CDP endpoint.
4. Through CDP, the script evaluates a small JavaScript snippet inside the panel that:
   - inserts a `<style>` element with RTL CSS,
   - installs a `MutationObserver` that replaces em-dash/en-dash with hyphen and arrows (U+2190 through U+2199, U+21D0 through U+21D4) with comma in visible text nodes,
   - skips `<pre>`, `<code>`, `<kbd>`, `<samp>` subtrees so source code is never modified.
5. A polling loop repeats every 2 seconds to handle panel reloads.
6. `scripts/tray-icon.ps1` runs in parallel as a zero-dependency PowerShell tray indicator (green / red / gray) driven by a status file at `%TEMP%\claude-word-rtl.status`.

## Persistent state (and where)

- **`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Claude for Word RTL Tray.lnk`** - per-user Startup-folder shortcut written by `install.bat`. Launches the tray icon at login via `wscript.exe scripts\start-tray.vbs`. Removed by `uninstall.bat`.
- **`HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL`** - per-user "Apps and Features" registration written by `install.bat` so the tool appears in Windows Settings > Apps > Installed apps as "Claude for Word RTL Fix". Values: `DisplayName`, `DisplayVersion`, `Publisher`, `InstallLocation`, `UninstallString` (points at `uninstall.bat`), `DisplayIcon`, `URLInfoAbout`, `NoModify`, `NoRepair`. Removed by `uninstall.bat`. HKLM is never touched; no other registry keys are written.
- **`%TEMP%\claude-word-rtl.status`**, **`.pid`**, **`.tray.pid`**, **`.lock`** - small transient files used for IPC between the injector, the tray, and the wrapper. All written at runtime only. Cleared by `cleanup.bat` and `uninstall.bat`.
- **`scripts/node_modules/`** - installed on first run. Removed by `uninstall.bat`.
- **`install.log`** - next to `install.bat`, captures output of the last install run. Covered by `.gitignore`. Delete manually if desired.

No file-association changes, no Scheduled Tasks, no system-wide services, no changes to `Normal.dotm` or any Word template.

If you are upgrading from v0.3.x, `uninstall.bat` also performs best-effort cleanup of legacy artifacts: the old "Word (RTL)" Start Menu shortcut and the old `HKCU\Software\Classes\Word.Document.*` overrides + `HKCU\Software\ClaudeWordRTL` marker.

## Processes and files loaded

- `cmd.exe` running one of the `.bat` scripts (all are plain-text and readable).
- `wscript.exe` running `inject-hidden.vbs` (one-line launcher for the Node process).
- `node.exe` running `scripts/inject.js`. Its PID is written to `%TEMP%\claude-word-rtl.pid` so cleanup scripts target this specific process and never mass-kill all `node.exe` on the machine.
- `powershell.exe` running `scripts/tray-icon.ps1` (uses `System.Windows.Forms.NotifyIcon` only).
- `scripts/node_modules/chrome-remote-interface` and its transitive dependencies, fetched from the public npm registry on first run.

No other binaries are installed or run by this tool.

## Risks

### Local debug port exposure

While Word is running via this tool, port `9222` listens on `localhost` (not the network). Any local process running under the same user account can connect to it and:

- read the DOM of the Claude panel, including visible conversation content,
- read cookies scoped to `claude.ai`,
- execute arbitrary JavaScript inside the panel.

This is the single most important risk. It is equivalent to running any Chromium-based application with `--remote-debugging-port` enabled and is a known local-attack surface.

**Mitigations:**
- Close Word (via the tray's **Disconnect** item, or normally, or via `cleanup.bat`) when you are not actively using the Claude panel.
- Do not run untrusted software on the same user account while Word is open via this tool.
- If you want to pause the tool temporarily, right-click the tray icon and pick **Exit**. Word then opens normally (no wrapper, no debug port) until you relaunch the tray.
- **Not exposed to the network** - WebView2 only binds to `127.0.0.1`. Remote hosts cannot reach it.

### Supply-chain trust in npm dependencies

`chrome-remote-interface` and its transitive dependencies are pulled from the public npm registry. A compromise of any of these packages could theoretically affect your machine. This is true of any Node.js project.

**Mitigations:** audit `scripts/package-lock.json` after the first install; pin versions; run `npm audit`.

### Unsigned scripts

The `.bat`, `.vbs`, `.ps1`, and `.js` files in this repo are not code-signed. Windows SmartScreen will warn on first run. Device Guard / WDAC on corporate-managed machines may block execution outright. Users who clone the repo via git or download the zip are expected to review the code before running - the surface is small enough to do so in about ten minutes.

### Enterprise environments

Organizations with DLP, EDR, or Group Policy controls may detect WebView2 debug-port enablement on Office as anomalous behavior. This tool is not suitable for sealed corporate laptops without coordination with your IT/security team.

## What the tool does not do

- Does not open any outbound network connection from its own code. The only network activity is:
  - `npm install` on first run (one-time, fetches dependencies from the public registry),
  - `check-update.bat` (optional, user-invoked, hits the GitHub releases API only),
  - normal Claude and Office traffic, unchanged, routed by Word itself.
- Does not read, write, or exfiltrate conversation content. The injected JavaScript modifies the DOM in place; nothing is serialized, transmitted, or written to disk by the injector beyond the one-line status/PID files listed above.
- Does not modify files outside its own folder and the `%TEMP%` transients listed above (plus the single per-user Startup-folder shortcut and the single per-user Apps-and-Features `Uninstall` registry key, both written by `install.bat` and removed by `uninstall.bat`).
- Does not change Word's file associations.
- Does not create Scheduled Tasks or services.
- Does not modify `Normal.dotm` or any Word template.
- Does not bypass, disable, or interfere with any guardrails, rate limits, or policies imposed by Claude or Office.

## Reviewing the code

The core is short enough to audit in one sitting:

- `scripts/inject.js` - CDP connection, CSS and observer script, polling loop (~340 lines).
- `scripts/tray-icon.ps1` - NotifyIcon indicator, Connect/Disconnect state machine, Auto-enable toggle (~630 lines).
- `scripts/check-update.js` - GitHub releases check (~105 lines, uses Node built-in `https` only).
- `install.bat`, `uninstall.bat`, `word-wrapper.bat`, `doctor.bat`, `cleanup.bat`, `start.bat`, `inject-hidden.vbs`, `scripts/start-tray.vbs` - plain-text launchers.

Everything else is documentation.
