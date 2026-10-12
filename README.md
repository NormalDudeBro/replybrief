# ReplyBrief

A standalone Windows companion for turning the conversation you are reading into a bullet-point
summary and suggested reply points. You write and send the response yourself.

**Status: developer prototype, not a finished multi-service product.** The Windows app and Outlook
connector are implemented. Gmail and Teams connectors are planned, not connected.

## Available now

- Windows desktop app with editable capture preview and source/context labels.
- Four analysis sections: Summary, Needs Attention, Suggested Reply Points, Missing Information.
- **Ollama only**, running on your workstation or a server on your network. No cloud AI provider or API key.
- Configurable server address and model, a connection check that lists installed models, cancellation,
  and a two-minute analysis timeout.
- Outlook web add-in that captures the selected message's subject, sender, and text body.
- A paired, loopback-only HTTPS handoff from the Outlook add-in to the desktop app.
- Capture and analysis are separate actions. Incoming captures never automatically call Ollama.
- A capture arriving during analysis waits for you to explicitly load it. Only the latest pending
  capture is retained; this is not a durable inbox.
- No conversation history, credentials, or model output written to disk by ReplyBrief. Text stays in
  memory until cleared or the app exits. Neither this claim nor Clear Capture guarantees forensic
  erasure from OS memory, crash dumps, or the clipboard.

The Outlook connector reads **one selected message**, including quoted text already in its body.
It does not retrieve separate thread messages, attachments, images, or linked files.

## Run

### Downloaded build

Download the Windows x64 ZIP from [Releases](https://github.com/NormalDudeBro/replybrief/releases),
extract the whole folder, and launch `ReplyBrief.Desktop.exe`. The build includes its .NET runtimes.
Enter the LAN Ollama address and model in the app. Keep all runtime files and the `outlook` folder
together. This is an unsigned portable prototype, not an installer; Windows may show an
unknown-publisher/SmartScreen warning. See [release notes](docs/RELEASE-NOTES.md).

Outlook's development-certificate setup still requires the .NET SDK; manual paste/analysis does not.
There is no automatic server discovery yet.

### From source

Prerequisites: Windows, .NET 8 or newer SDK, and an Ollama server with a pulled model.

From this directory in PowerShell:

```powershell
# Substitute your real LAN server and installed model.
$env:REPLYBRIEF_OLLAMA_URL = "http://192.168.1.50:11434"
$env:REPLYBRIEF_OLLAMA_MODEL = "llama3.1:8b"
dotnet run --project src\ReplyBrief.Desktop\ReplyBrief.Desktop.csproj
```

Alternatively, edit the server address and model in the app. Defaults are `http://localhost:11434`
and `llama3.1`. Settings are session-only; the two environment variables provide repeatable startup
defaults. The app does not silently save them.

Use **Check connection**, paste a conversation to try the desktop workflow, review the destination
and consent checkbox, then click **Analyze conversation**. Check Connection fetches model names only,
not conversation text. Clearing or editing a capture clears old results and resets consent.

To use New Outlook or Outlook on the web, follow [Outlook setup](docs/outlook.md). You must explicitly
trust the local development HTTPS certificate and sideload the add-in; your organization may block
custom add-ins. The app does not change your certificate trust store automatically.

## LAN Ollama

On the server, pull the intended model and configure Ollama to listen on its LAN interface. Ollama's
default loopback binding is not reachable from other computers. Restart its service after changing
`OLLAMA_HOST`; use an interface-specific address when possible. Restrict inbound TCP 11434 with the
server firewall to the clients that need it. Do **not** forward the port from your router or expose
an unauthenticated Ollama endpoint to the internet.

The desktop app calls the server directly, so browser CORS settings such as `OLLAMA_ORIGINS` are not
required for analysis. The Outlook add-in calls only the local capture host, not the LAN model server.

Plain HTTP sends conversation text unencrypted across the network. Prefer an HTTPS reverse proxy
with a certificate trusted by Windows when the network is not fully trusted. ReplyBrief does not
bypass certificate checks, follow HTTP redirects, or support proxy authentication yet. It accepts
configured HTTP/HTTPS hosts without proving they are on your LAN; verify the destination yourself.
The Ollama server and any proxy may retain data or logs independently of ReplyBrief.

The analyzer requests Ollama's JSON-schema structured output, `stream: false`, a 16,384-token context,
and up to 2,048 output tokens. Use a recent Ollama version and an instruction model that supports
structured output. Available RAM/VRAM, context limits, and model quality affect results. The 60,000
character capture limit is a safety bound, **not** a guarantee that every model can read the whole
capture. Shorten long messages when needed. Cancellation stops waiting and closes the request; it
does not guarantee that the remote server immediately stops inference.

## Safety

Conversation content and metadata are treated as untrusted input in the system prompt. There are no
model tools or send-message actions. Output must contain all four correctly typed, bounded sections;
malformed, truncated, and oversized responses fail instead of becoming apparently successful notes.
Prompt instructions reduce risk but do not make a model immune to prompt injection or hallucination.
Review every suggested point before using it, especially facts, dates, commitments, and sensitive data.

## Build and tests

```powershell
dotnet build ReplyBrief.sln -c Release
dotnet test ReplyBrief.sln -c Release
bun test tests/outlook
bun x office-addin-manifest validate integrations/outlook/manifest.xml
```

The .NET tests exercise Core over a loopback Ollama protocol fixture, the real capture host over HTTPS,
and desktop view-model behavior on an STA dispatcher. Connector tests need an existing .NET
development certificate (`dotnet dev-certs https`), but do not require trust or change the trust store.
The JavaScript tests use a simulated Office runtime; they do not replace testing a real Outlook client.

Create a self-contained Windows x64 release ZIP and SHA256 checksum with:

```powershell
.\scripts\Publish-Windows.ps1 -Version 0.1.0-preview.1
```

Artifacts go to the ignored `artifacts` directory. The script refuses to overwrite an existing
version's package. Review and test artifacts before publishing them.

Verified: builds, automated protocol/security/UI-state tests, and Microsoft's manifest schema validator.
**Not yet verified:** sideloading in a real New Outlook/Microsoft 365 account, analysis with your LAN
Ollama server, Gmail capture, Teams capture, installation on a clean Windows machine, or organization-wide deployment.

## Structure

| Directory | Responsibility |
| --- | --- |
| `src/ReplyBrief.Core` | Capture validation, Ollama requests, structured result validation |
| `src/ReplyBrief.Connectors` | Paired local HTTPS capture endpoint and Outlook asset hosting |
| `src/ReplyBrief.Desktop` | WPF preview, settings, analysis, cancellation, and result display |
| `integrations/outlook` | Outlook manifest, task pane, and help assets |
| `tests` | .NET and Outlook JavaScript regression tests |

See [the product plan](docs/PLAN.md) for remaining integration work and acceptance criteria.
