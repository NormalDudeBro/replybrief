# v0.1.0-preview.1

First developer preview of ReplyBrief, a standalone Windows conversation companion using an Ollama
server on your network. This is a prototype, not a finished multi-service email/chat product.

## Download and run

1. Download `ReplyBrief-0.1.0-preview.1-win-x64.zip` and its `.sha256` checksum.
2. Extract the complete ZIP. Keep the executable, runtime files, and `outlook` folder together.
3. Run `ReplyBrief.Desktop.exe` inside the extracted folder.
4. Enter your LAN Ollama URL and installed model. Use **Check connection**.
5. Paste a conversation, review it, confirm the server destination, and analyze it.

The Windows x64 build bundles .NET 8.0.31: no separate .NET runtime installation is needed. It is
unsigned, so Windows may display an unknown-publisher/SmartScreen warning. Only download from this
repository's Releases page. The executable is a portable app, not an installer.

## Features

- Editable conversation preview and context labels.
- Bullet-point Summary, Needs Attention, Suggested Reply Points, and Missing Information.
- Ollama-only analysis with a configurable network server and model.
- Connection/model check, cancellation, timeout, and structured response validation.
- Outlook selected-message capture through an add-in and paired loopback HTTPS host.
- Explicit analysis; no automatic sending or conversation persistence by ReplyBrief.

## Limits and setup

- **Gmail and Teams are planned, not implemented.** Manual paste is available for any source.
- Outlook captures the selected message only, not separate thread messages or attachments.
- Outlook setup requires a trusted localhost HTTPS certificate and custom add-in installation.
  The current certificate setup command, `dotnet dev-certs https --trust`, requires the .NET SDK.
  ReplyBrief never changes your certificate trust store automatically. If no development certificate
  exists, Outlook capture is unavailable but the manual desktop workflow still works.
- Custom add-ins may be blocked by your organization. See `docs/outlook.md` in the ZIP.
- Real New Outlook sideloading and analysis through the user's LAN server remain unverified.
- No installer, code signing, saved settings, authenticated Ollama proxy support, or automatic updates.
- HTTP sends content unencrypted across the LAN. Use a trusted network or a trusted HTTPS proxy, and
  never expose an unauthenticated Ollama server to the internet.
- Model results need human review. Prompt-injection resistance and factual accuracy are not guaranteed.

## Verification

- Release build: zero warnings and errors.
- 186 .NET tests passed.
- 19 Outlook JavaScript tests passed.
- Outlook manifest passed Microsoft's schema validator.

Automated tests use local protocol fixtures and a simulated Office runtime; they are not end-to-end
verification of real Outlook or a real language model.
