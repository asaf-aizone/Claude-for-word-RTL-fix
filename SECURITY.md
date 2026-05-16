# Security Policy

## Reporting a vulnerability

If you find a security issue (e.g., a way the tool can be misused to
leak data, or a flaw in the CSS/observer injection), please do **not**
file a public GitHub issue.

Contact: open a GitHub Security Advisory on this repository, or contact
the maintainer privately via [LinkedIn](https://www.linkedin.com/in/asaf-abramzon-7a2b61180/)
or GitHub ([@asaf-aizone](https://github.com/asaf-aizone)).

Expected response time: best effort, typically within a week.

## Supported versions

Only the `main` branch is supported. Older tags are not patched.

## Security model - short version

While an Office app (Word, Excel, PowerPoint, or Outlook) is running
through this tool, that app's WebView2 host exposes a Chrome DevTools
debug port on a dynamic localhost port number (one port per Office
WebView2 host process; v0.2.0 uses `--remote-debugging-port=0`, was
a fixed `9222` in v0.1.x). The port is:

- **Local-only** - not reachable over the network.
- **Live only while the Office app is running** - closes when you close
  Word, Excel, PowerPoint, or Outlook.
- **Unauthenticated** - any other process running under your user
  account can connect to it and read the Claude panel's DOM, observe
  its network traffic, or inject JavaScript into it.

This is inherent to enabling WebView2 debugging; it is the same exposure
as running any Chromium-based application with `--remote-debugging-port`.
The flag is set in the wrapper's process scope only and is inherited
just by the launched Office app, not by Teams, the standalone Outlook
reading pane outside of this tool's wrapper, Edge, or any other
WebView2 host on the account.

### Outlook is opt-in with a stricter model (v0.3.0+)

Outlook is treated separately because mail content enters the Claude
panel DOM during "Summarize this email" / "Draft a reply", and the
exposure is qualitatively more sensitive than for Word/Excel/PowerPoint
document panels. To address that, v0.3.0 adds:

- The injector permanently blocks Outlook unless a per-launch opt-in
  flag (written only by `outlook-wrapper.bat` after the user clicks
  Connect Outlook and OKs a warning dialog whose default-focused
  button is Cancel) is present. The flag is cleared at every injector
  startup so no consent carries silently across sessions.
- No auto-launch: the tray will not bring the injector up just because
  Outlook is running. Word/Excel/PowerPoint do trigger auto-launch;
  Outlook does not.
- 15-minute auto-disconnect timer per Outlook attachment, after which
  the injector revokes the opt-in flag on its own.
- A dedicated **Disconnect Outlook only** tray item to drop just the
  Outlook attachment without affecting any other connected app.
- Tenant-correlated URL parameters are redacted from the diagnostic
  log for Outlook only.

New Outlook (`olk.exe`) is intentionally out of scope; the wrapper
and the Connect Outlook flow both refuse to run if it is detected.

Full design rationale and the M0 finding that drove the opt-in model:
[docs/security.md - Outlook-specific risks and mitigations](docs/security.md#outlook-specific-risks-and-mitigations).

Recommendations:

- Close the Office app when you are not actively using the Claude panel.
- For Outlook: prefer **Disconnect Outlook only** between mail sessions
  instead of relying on the 15-minute timer.
- Do not install untrusted browser extensions or other untrusted
  software on the same user account.
- On corporate-managed machines with EDR/DLP agents, check with your
  IT/security team before enabling a debug port on Office.

For the full threat model (what a local attacker can and cannot do,
what the tool does and does not touch), see [docs/security.md](docs/security.md).

## What this tool does and does not do

- Does not open outbound network connections from its own code.
- Does not store, transmit, or log conversation content.
- Does not modify Office file associations.
- Does not create scheduled tasks or services.
- Does not modify `Normal.dotm`, any Word template, any Excel
  workbook template, any PowerPoint template, or any Outlook
  signature/account/rule.
- `install.bat` creates one per-user Startup-folder shortcut
  (`Claude for Word RTL Tray.lnk` - filename retained for v0.1.x
  upgrade compatibility, even though v0.3.0 covers Word, Excel,
  PowerPoint, and Outlook) that launches the tray icon at login,
  and one per-user
  `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL`
  registry key so the tool appears in Windows Settings > Apps. The
  `DisplayName` in that key is still "Claude for Word RTL Fix" for the
  same upgrade-compatibility reason; v0.2.0 deliberately did not rename
  it so reinstalls over v0.1.x replace the existing entry instead of
  duplicating it. No other registry writes. No admin required. Both
  are reversed by `uninstall.bat`.

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

The authors provide no warranty; use at your own risk.
