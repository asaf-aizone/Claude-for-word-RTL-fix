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

While Word is running with this tool, WebView2 exposes a Chrome DevTools
debug port on `localhost:9222`. The port is:

- **Local-only** - not reachable over the network.
- **Live only while Word is running** - closes when you close Word.
- **Unauthenticated** - any other process running under your user
  account can connect to it and read the Claude panel's DOM, observe
  its network traffic, or inject JavaScript into it.

This is inherent to enabling WebView2 debugging; it is the same exposure
as running any Chromium-based application with `--remote-debugging-port`.

Recommendations:

- Close Word when you are not actively using the Claude panel.
- Do not install untrusted browser extensions or other untrusted
  software on the same user account.
- On corporate-managed machines with EDR/DLP agents, check with your
  IT/security team before enabling a debug port on Office.

For the full threat model (what a local attacker can and cannot do,
what the tool does and does not touch), see [docs/security.md](docs/security.md).

## What this tool does and does not do

- Does not open outbound network connections from its own code.
- Does not store, transmit, or log conversation content.
- Does not modify Word's file associations.
- Does not create scheduled tasks or services.
- Does not modify `Normal.dotm` or any Word template.
- `install.bat` creates one per-user Startup-folder shortcut
  (`Claude for Word RTL Tray.lnk`) that launches the tray icon at login,
  and one per-user `HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClaudeWordRTL`
  registry key so the tool appears in Windows Settings > Apps. No other
  registry writes. No admin required. Both are reversed by
  `uninstall.bat`.

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
