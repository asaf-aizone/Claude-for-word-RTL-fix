# Security and Threat Model

This document explains what the tool does at runtime, what risks exist, and what it specifically does not do. Read this before using in a sensitive environment.

## Entry points

The tool can be started in two ways, both of which end up with the same set of running processes (one or more of Word/Excel/PowerPoint, plus a single hidden Node injector serving all three):

1. **Tray icon (recommended).** After `install.bat` runs, a per-user Startup-folder shortcut (`Claude for Word RTL Tray.lnk`, filename retained for v0.1.x upgrade compatibility) launches `scripts/start-tray.vbs` at login, which launches the tray icon. The user opens Word, Excel, or PowerPoint normally and then right-clicks the tray icon and picks **Connect Word**, **Connect Excel**, or **Connect PowerPoint**; the tray enumerates the app's open documents/workbooks/presentations via COM, gracefully closes the app, and relaunches it through the matching wrapper (`word-wrapper.bat`, `excel-wrapper.bat`, or `powerpoint-wrapper.bat`) with the WebView2 debug flag set and the same files reopened. No registry writes, no file-association changes. Fully reversed by `uninstall.bat`.
2. **`start.bat`** - debug mode for Word. Launches Word and the injector in the foreground with a visible log window. Useful for troubleshooting. Leaves no persistent state behind. Excel and PowerPoint do not have an equivalent debug-mode launcher; use the tray's Connect items.

## What happens at runtime

1. The relevant wrapper (`word-wrapper.bat`, `excel-wrapper.bat`, or `powerpoint-wrapper.bat`) sets the environment variable `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=0` in its own process scope. This is a Microsoft-documented WebView2 flag used for add-in debugging. Port=0 means WebView2 picks a free dynamic port per process (was a fixed `9222` in v0.1.x); `start.bat` keeps the same per-process pattern.
2. The wrapper launches the Office app. The app's WebView2 host inherits the flag and exposes Chrome DevTools Protocol on whichever ephemeral localhost port the OS allocated. Multiple Office apps launched through their respective wrappers each get their own debug surface; in some cases (notably Word + PowerPoint) two apps share a single WebView2 host process and therefore the same port, with multiple page targets distinguished by the `_host_Info=` URL parameter.
3. `scripts/inject.js` runs (hidden via `inject-hidden.vbs`, or visibly via `start.bat`). It uses `scripts/port-discovery.js` each tick to enumerate active CDP ports: walk `tasklist` for `msedgewebview2.exe` PIDs, map each PID to its LISTENING port via `netstat`, probe each candidate's `/json/list`, and collect every Claude target. For each target it opens a WebSocket to that target's CDP endpoint. App identification per target reads the `_host_Info=` URL parameter that Office appends to the panel URL.
4. Through CDP, the script evaluates a small JavaScript snippet inside the panel that:
   - inserts a `<style>` element with RTL CSS,
   - installs a `MutationObserver` that replaces em-dash/en-dash with hyphen and arrows (U+2190 through U+2199, U+21D0 through U+21D4) with comma in visible text nodes,
   - skips `<pre>`, `<code>`, `<kbd>`, `<samp>` subtrees so source code is never modified.
5. A polling loop repeats every 2 seconds to handle panel reloads and to pick up newly-launched Office apps.
6. `scripts/tray-icon.ps1` runs in parallel as a zero-dependency PowerShell tray indicator (green / red / gray) driven by the aggregate status file at `%TEMP%\claude-word-rtl.status` (icon color) and the per-app status file at `%TEMP%\claude-office-rtl.apps.json` (per-app status labels at the top of the menu).

## Persistent state (and where)

- **`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Claude for Word RTL Tray.lnk`** - per-user Startup-folder shortcut written by `install.bat`. Launches the tray icon at login via `wscript.exe scripts\start-tray.vbs`. Removed by `uninstall.bat`.
- **`HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL`** - per-user "Apps and Features" registration written by `install.bat` so the tool appears in Windows Settings > Apps > Installed apps as "Claude for Word RTL Fix". Values: `DisplayName`, `DisplayVersion`, `Publisher`, `InstallLocation`, `UninstallString` (points at `uninstall.bat`), `DisplayIcon`, `URLInfoAbout`, `NoModify`, `NoRepair`. Removed by `uninstall.bat`. HKLM is never touched; no other registry keys are written.
- **`%TEMP%\claude-word-rtl.status`**, **`.pid`**, **`.tray.pid`**, **`.lock`**, plus **`%TEMP%\claude-office-rtl.apps.json`** (v0.2.0+ per-app state file written by the injector and read by the tray for the three per-app status labels) - small transient files used for IPC between the injector, the tray, and the three wrappers. All written at runtime only. Cleared by `cleanup.bat` and `uninstall.bat`. The `claude-word-rtl.*` filenames are kept for v0.1.x upgrade compatibility even though the tool now covers all three Office apps; the prefix is a name, not a constraint.
- **`scripts/node_modules/`** - installed on first run. Removed by `uninstall.bat`.
- **`install.log`** - next to `install.bat`, captures output of the last install run. Covered by `.gitignore`. Delete manually if desired.

No file-association changes, no Scheduled Tasks, no system-wide services, no changes to `Normal.dotm` or any Word/Excel/PowerPoint template.

If you are upgrading from v0.1.x, `uninstall.bat` also performs best-effort cleanup of legacy artifacts: the old "Word (RTL)" Start Menu shortcut and the old `HKCU\Software\Classes\Word.Document.*` overrides + `HKCU\Software\ClaudeWordRTL` marker. v0.1.x also wrote `HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` when Auto-enable was on; v0.2.0 clears it on install and uninstall, only when its value matches one of our known strings (`=9222` from v0.1.x or `=0` from a v0.2.0 prerelease).

## Processes and files loaded

- `cmd.exe` running one of the `.bat` scripts (all are plain-text and readable).
- `wscript.exe` running `inject-hidden.vbs` (one-line launcher for the Node process).
- `node.exe` running `scripts/inject.js`. Its PID is written to `%TEMP%\claude-word-rtl.pid` so cleanup scripts target this specific process and never mass-kill all `node.exe` on the machine.
- `powershell.exe` running `scripts/tray-icon.ps1` (uses `System.Windows.Forms.NotifyIcon` only).
- `scripts/node_modules/chrome-remote-interface` and its transitive dependencies, fetched from the public npm registry on first run.

No other binaries are installed or run by this tool.

## Risks

### Local debug port exposure

While Word, Excel, or PowerPoint is running via this tool, that app's WebView2 host listens on a dynamic localhost port (one per Office WebView2 host process; the OS picks a free port via `--remote-debugging-port=0`). The port is bound to localhost only, not the network. Any local process running under the same user account can connect to it and:

- read the DOM of the Claude panel, including visible conversation content,
- read cookies scoped to `claude.ai`,
- execute arbitrary JavaScript inside the panel.

This is the single most important risk. It is equivalent to running any Chromium-based application with `--remote-debugging-port` enabled and is a known local-attack surface. Dynamic ports do not change the threat model versus the v0.1.x fixed-9222 model: a local attacker can enumerate listening ports just as easily as it could connect to a known port number, and `port-discovery.js` itself is the proof of how a friendly process does so.

**Mitigations:**
- Close the Office app (via the tray's **Disconnect all** item, or normally, or via `cleanup.bat`) when you are not actively using the Claude panel.
- Do not run untrusted software on the same user account while an Office app is open via this tool.
- If you want to pause the tool temporarily, right-click the tray icon and pick **Exit**. Office apps then open normally (no wrapper, no debug port) until you relaunch the tray.
- **Not exposed to the network** - WebView2 only binds to `127.0.0.1`. Remote hosts cannot reach it.

### Supply-chain trust in npm dependencies

`chrome-remote-interface` and its transitive dependencies are pulled from the public npm registry. A compromise of any of these packages could theoretically affect your machine. This is true of any Node.js project.

**Mitigations:** audit `scripts/package-lock.json` after the first install; pin versions; run `npm audit`.

### Unsigned scripts

The `.bat`, `.vbs`, `.ps1`, and `.js` files in this repo are not code-signed. Windows SmartScreen will warn on first run. Device Guard / WDAC on corporate-managed machines may block execution outright. Users who clone the repo via git or download the zip are expected to review the code before running - the surface is small enough to do so in about ten minutes.

### Enterprise environments

Organizations with DLP, EDR, or Group Policy controls may detect WebView2 debug-port enablement on Office as anomalous behavior. This tool is not suitable for sealed corporate laptops without coordination with your IT/security team.

## Outlook-specific risks and mitigations

Outlook is supported (v0.3.0+) under a stricter security model than Word, Excel and PowerPoint. The reason: when the user asks the Claude add-in to summarize an email or draft a reply, the **content of the active email** becomes part of the panel DOM for the duration of that operation. The same CDP attach surface that lets the injector apply RTL CSS also lets any local process under the same user account read that DOM. For a Word document panel the comparable exposure is content the user explicitly pasted or pinned. For Outlook the exposure is mail itself, which can include credentials, MFA codes, legal documents, account recovery links, and tenant identifiers. The exposure window is short (only while a summarize/draft is in flight), but the content class is qualitatively more sensitive.

Because of this, Outlook has the following additional protections that **do not** apply to Word/Excel/PowerPoint:

- **No silent attach.** The injector maintains a per-app block list (`lib/office-apps.js` `BLOCKED_HOST_INFO_KEYS`). Outlook is permanently on it. Discovery still finds the Outlook CDP target each tick, but the injector logs it as blocked and does not attach. Attach happens only when a per-launch opt-in flag file (`%TEMP%\claude-office-rtl.outlook-optin`) exists. The flag is **never persisted across injector restarts** - the injector clears it at startup.
- **No auto-launch.** When any Word/Excel/PowerPoint app is up but the injector is gone (crash, manual kill), the tray relaunches the injector automatically. Outlook is excluded from this trigger - if Outlook is the only Office app running, the injector stays down. The user must explicitly pick **Connect Outlook** to start a session.
- **Per-launch warning dialog.** **Connect Outlook** shows a OK/Cancel dialog whose default focused button is Cancel, naming the exact exposure ("local process can read panel DOM while attached") rather than hand-waving about "security". A stray Enter press does not opt the user in.
- **Strict target filter.** The injector validates the CDP target's `_host_Info=` URL parameter before attaching: it URL-decodes the value, takes the first `$`-delimited segment, lowercases it, and compares it for equality against `outlook`. A target whose `_host_Info=` does not start with an `Outlook`-family token is not classified as Outlook and is not gated by the Outlook opt-in. Note that the comparison is case-insensitive on the first segment, so this filter does not by itself defeat a local process that constructs its own WebView2 with a forged `_host_Info=`; defence-in-depth for that scenario relies on the opt-in flag and the auto-disconnect timer (above), not on the filter alone.
- **Auto-disconnect after 15 minutes.** Even if the user forgets to disconnect, the injector tears the Outlook CDP client down after 15 minutes of continuous attachment and revokes the opt-in flag. The user must click **Connect Outlook** again to start a new session. Tunable in source via `OUTLOOK_AUTO_DISCONNECT_MIN` in `scripts/inject.js`.
- **URL redaction in the diagnostic log.** The Office add-in URL contains an `et=` parameter holding base64-encoded tenant metadata (account id, tenant id, expiry). For Outlook only, the injector strips every query parameter from logged URLs except `_host_Info=` and replaces the rest with `[redacted]`. The log still shows enough to diagnose connection problems but no longer leaks tenant identifiers to anyone with read access to `%TEMP%`.
- **Disconnect Outlook only.** The tray menu item **Disconnect Outlook only** drops the Outlook CDP attachment without closing Outlook itself and without affecting Word/Excel/PowerPoint sessions. Use it when you are done with mail but still working in another Office app. Implementation: the tray writes a request file the injector polls each tick; the injector closes the Outlook CDP client(s), clears the auto-disconnect timer, and revokes the opt-in flag.

**Recommendations for Outlook users:**

- Keep Outlook closed when not in active use, or use **Disconnect Outlook only** between mail sessions instead of relying on the 15-minute timer.
- Treat the warning dialog as a real consent gate. Click Cancel if you are not currently planning to use Claude on a specific email.
- The 15-minute auto-disconnect is a backstop, not a budget. There is no business case for an 8-hour mail day with the CDP port continuously open.
- New Outlook (`olk.exe`) is intentionally not supported. The probe in M0 deferred it, and the **Connect Outlook** flow refuses to launch if New Outlook is detected to avoid colliding on shared per-user state. If you have both installed, close New Outlook before connecting.

For the full design rationale and the M0 evidence that motivated the opt-in model, see `docs/OUTLOOK-EXPANSION-PLAN.md` sections 3 and 4 and `probe/README.md` "silent CDP attach".

## What the tool does not do

- Does not open any outbound network connection from its own code. The only network activity is:
  - `npm install` on first run (one-time, fetches dependencies from the public registry),
  - `check-update.bat` (optional, user-invoked, hits the GitHub releases API only),
  - normal Claude and Office traffic, unchanged, routed by Word itself.
- Does not read, write, or exfiltrate conversation content. The injected JavaScript modifies the DOM in place; nothing is serialized, transmitted, or written to disk by the injector beyond the one-line status/PID files listed above.
- Does not modify files outside its own folder and the `%TEMP%` transients listed above (plus the single per-user Startup-folder shortcut and the single per-user Apps-and-Features `Uninstall` registry key, both written by `install.bat` and removed by `uninstall.bat`).
- Does not change Office file associations.
- Does not create Scheduled Tasks or services.
- Does not modify `Normal.dotm` or any Word/Excel/PowerPoint template.
- Does not bypass, disable, or interfere with any guardrails, rate limits, or policies imposed by Claude or Office.

## Reviewing the code

The core is short enough to audit in one sitting:

- `scripts/inject.js` - CDP connection, CSS and observer script, polling loop, per-app status writer.
- `scripts/port-discovery.js` - dynamic-port enumeration via `tasklist` + `netstat` + `/json/list` probe; per-target app identification via `_host_Info=`.
- `lib/office-apps.js` - shared per-app metadata (process name, `_host_Info` key) used by the injector and the probes.
- `scripts/tray-icon.ps1` - NotifyIcon indicator, per-app Connect state machines, Disconnect-all, per-app status labels.
- `scripts/check-update.js` - GitHub releases check (~105 lines, uses Node built-in `https` only).
- `install.bat`, `uninstall.bat`, `word-wrapper.bat`, `excel-wrapper.bat`, `powerpoint-wrapper.bat`, `doctor.bat`, `cleanup.bat`, `start.bat`, `inject-hidden.vbs`, `scripts/start-tray.vbs` - plain-text launchers.

Everything else is documentation.

## Anthropic Terms of Service compliance

This tool operates entirely on the user's machine. It does not access
Anthropic's API, does not reverse-engineer the Service, does not bypass
guardrails, rate limits, or safety systems, and does not scrape or
harvest data. The conversation between the user and Claude is unchanged
by this tool. The Claude add-in receives the same input and produces
the same output it would have produced without this tool installed.

What the tool does is restyle the locally-rendered HTML inside
Microsoft's WebView2 process on the user's own machine, adding RTL CSS
rules and replacing certain glyphs (em-dash, en-dash, several arrow
characters) in already-rendered text. Functionally this is equivalent
to a user stylesheet, a browser accessibility extension, or a screen
reader: the underlying Service is untouched; only the local rendering
is adapted so a Hebrew-speaking user can read it.

Anthropic's Acceptable Use Policy and Consumer Terms govern the user's
account at all times, regardless of whether this tool is installed.
The user is responsible for compliance with those terms. If Anthropic's
terms change to restrict client-side modifications, the user should
comply with Anthropic's terms over this tool.

The tool is open source under Apache 2.0. The full source - injector,
tray, installer, every script - is auditable in this repository. There
is no obfuscation, no compiled binary, no closed-source component.
