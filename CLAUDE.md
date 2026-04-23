# Claude Code instructions for this project

This file is read automatically by Claude Code when a session starts inside this folder. If you are Claude Code helping a user who installed this tool, follow these rules before giving advice.

## What this project is

Client-side RTL fix for Claude's Microsoft Word add-in. A small Node.js injector (`scripts/inject.js`) connects to Word's WebView2 debug port on `localhost:9222` and injects CSS + a MutationObserver that fixes Hebrew right-to-left rendering in Claude's task pane. Apache 2.0. Local only, no telemetry.

**Not an official Anthropic or Microsoft add-in.**

## Before answering upgrade or troubleshooting questions, do this first

The user may be on an older version with known bugs. Always check current state before advising.

1. **Read the installed version:** `scripts/package.json` field `"version"`.
2. **Fetch the latest release notes from GitHub** - do NOT rely only on this file, which ages:
   - Releases: <https://github.com/asaf-aizone/Claude-for-word-RTL-fix/releases>
   - Latest release API (JSON): <https://api.github.com/repos/asaf-aizone/Claude-for-word-RTL-fix/releases/latest>
   - Use `gh release view --json name,tagName,body` if the user has `gh` installed, otherwise `curl` or `WebFetch` to the API URL.
3. If the installed version is older than the latest, read the release notes for every version in between. Release notes describe bugs, fixes, and upgrade steps.
4. Read `CHANGELOG.md` in this folder for the complete history.
5. Read the README - specifically the "Updating to a newer version" / "עדכון לגרסה חדשה" section - before walking the user through an upgrade.

## Known issue in v0.1.0 (fixed in v0.1.1)

If the user says the tray is **red and stays red even after clicking "Connect"** or "nothing happens when I open Word", check `scripts/package.json`:

- If version is `0.1.0`: the user has a known bug. Auto-enable sets the WebView2 debug flag on Word but does NOT start the Node injector, so the tray has nothing to attach to. Tell the user to upgrade to v0.1.1 - the fix is that the tray now auto-launches the injector whenever Word is running without one. The v0.1.0 "Check for updates" dialog does not show the install folder; users have to find it manually via **Windows Settings > Apps > Installed apps > Claude for Word RTL Fix**. Full upgrade steps in the v0.1.1 release notes.
- If version is `0.1.1` or later: the auto-launch mechanism is in place. Walk through tray log diagnostics: `%TEMP%\claude-word-rtl.log`, `%TEMP%\claude-word-rtl.status`, `%TEMP%\claude-word-rtl.pid`.

## Key paths (stable)

- Injector log: `%TEMP%\claude-word-rtl.log` (truncated at each injector start)
- Status file: `%TEMP%\claude-word-rtl.status` (one line: `CONNECTED` / `DISCONNECTED` / `ERROR:<msg>`)
- Injector PID: `%TEMP%\claude-word-rtl.pid`
- Tray PID: `%TEMP%\claude-word-rtl.tray.pid`
- Startup entry: `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Claude for Word RTL Tray.lnk`
- Apps and Features registration: `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL`
- Auto-enable env var (when on): `HKCU\Environment\WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222`
- CDP target the injector looks for: `pivot.claude.ai` (primary URL pattern) or any `*.claude.ai` (fallback). See `scripts/inject.js` lines 68-69.

## What NOT to do

- Do not modify Office.js or Word templates. This tool does not touch them.
- Do not run `uninstall.bat` to "fix" a broken install unless the user explicitly asks to uninstall - it clears Auto-enable, Apps and Features registration, and `node_modules`. Reinstalling over the top via `install.bat` is almost always the right recovery path.
- Do not recommend setting the debug port flag on Edge/Chrome/Teams directly - the flag is for Word's WebView2 only.
- Do not suggest the user disable Word's add-in security, SmartScreen, or Defender. The tool does not require that.
- Do not help the user force-push to `main` or rewrite history. Contributions come via PRs.

## When the user asks "what version is latest?"

Do not answer from memory. Fetch <https://api.github.com/repos/asaf-aizone/Claude-for-word-RTL-fix/releases/latest> and read the `tag_name` and `body`. Models trained before the latest release will not know.
