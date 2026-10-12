# Outlook add-in

The Outlook connector lets a person hand the message they are reading to the ReplyBrief desktop app
on the same machine. The add-in is a task pane that appears on the **Message** tab while a message is
open. It reads one message, shows you what it read, and only sends anything when you press the button.

Nothing is sent anywhere automatically, and the add-in itself never contacts an AI service.

## What is made of

| Piece | Where |
| --- | --- |
| Loopback capture host | `src/ReplyBrief.Connectors/LocalCaptureHost.cs` |
| Add-in sources (manifest, task pane, function file, support page, icons) | `integrations/outlook/` |
| Connector tests | `tests/ReplyBrief.Connectors.Tests/` |
| Task pane tests | `tests/outlook/taskpane.test.js` |

The desktop app references `ReplyBrief.Connectors`, which copies `integrations/outlook/*` to
`outlook/` next to its own output at build time and serves it from the build output at runtime. Edit
the files in `integrations/outlook`; they are the source of truth, and a rebuild picks the changes up.

## Using it from the desktop app

```csharp
await using var host = await LocalCaptureHost.StartAsync(capture =>
{
    // The desktop decides everything from here: queue it, analyse it, show it.
    Console.WriteLine($"{capture.Title} from {capture.Participants}");
});

Console.WriteLine($"Show this pairing code: {host.PairingCode}");
Console.WriteLine($"Add-in origin: {host.BaseAddress}");        // https://localhost:47831/
Console.WriteLine($"Manifest URL : {new Uri(host.BaseAddress, "/outlook/manifest.xml")}");
```

`StartAsync` also has a cancellation token, so a desktop app that is shutting down can stop the
listener deterministically.

The callback is invoked on a background pump, never on the request thread, so a slow analysis does not
make the task pane wait. The pump holds a **bounded** queue of 8 captures: if the desktop callback is
slower than the person clicking, the host answers `429 Too Many Requests` rather than buffering
without limit, and the task pane says ReplyBrief is busy. The host has no durable queue and does not
persist anything across runs, so if ReplyBrief is closed a queued capture is gone; anything that must
survive a restart belongs in the desktop app, which owns that decision.

`LocalCaptureHostOptions` is internal to the connector assembly. There is no supported way to change
the port, the asset directory or the dispatch-failure callback from outside; the public surface is
`StartAsync`, `PairingCode`, `BaseAddress`, `Port` and `DisposeAsync`.

## Security model

* **HTTPS only, loopback only.** Kestrel binds both loopback families on port `47831` with the ASP.NET
  Core development certificate, so `localhost` resolves whether the client lands on `127.0.0.1` or
  `::1`. There is no plain HTTP listener and no other interface is bound, so another machine on the
  network cannot reach it.
* **Pairing code.** Every run issues a new 43-character code (32 CSPRNG bytes, unpadded base64url).
  It travels in the `X-ReplyBrief-Pairing` request header and is compared with a length guard plus a
  fixed-time comparison. It never appears in a URL, in a served asset, in a response header or in a
  log.
* **Same origin.** If a request carries an `Origin` header it must be `https://localhost:47831`. The
  check compares against the port the connection actually landed on, not the `Host` header. No CORS
  policy is configured at all.
* **Bounded input.** Bodies are capped at 300 KB, JSON depth is capped at 8, only `application/json`
  is accepted, `source` must be `outlook`, and the payload is then re-validated by
  `ConversationCapture.Validate()` from Core.
* **Closed asset list.** Only `/outlook/taskpane.html`, `/outlook/taskpane.js`, `/outlook/commands.html`,
  `/outlook/support.html`, `/outlook/icon.svg`, `/outlook/icon-{16,32,80,128}.png` and
  `/outlook/manifest.xml` are served. Requests are resolved by exact route match against that list, so
  traversal attempts are 404s by construction. `commands.html` exists only because the manifest schema
  requires a function file; it loads office.js and nothing else, so the capture script can never run in
  a context with no UI and no pairing code.
* **Shutting down is visible.** Once disposal starts, new captures are refused with `503` rather than
  being accepted into a queue nobody will drain.
* **No CORS, no logging.** The host clears all configuration sources and all logging providers at
  startup, so an `appsettings.json`, an `ASPNETCORE_*` variable or a default logger cannot pull a
  capture body or a pairing code into a log file or add a second endpoint.

### About permissions and approvals

The manifest requests **`ReadItem`**, which is the minimum needed to read the item the user already
has open. Reading the current item is **not** an OAuth flow: there is no sign-in, no token and no
identity permission to grant for ReplyBrief's actual behaviour.

That does **not** mean installing it is frictionless. Sideloading a custom add-in is an installation
action that the person or their organisation has to allow:

* The mailbox typically needs the **My Custom Apps** role (or an equivalent Exchange assignment) for
  the person to install a custom add-in.
* An administrator can block custom add-ins outright through Exchange policies, and the manifest then
  simply never appears. ReplyBrief cannot override that and does not try to.
* Some organisations also require add-in installation to go through the admin centre rather than a
  developer sideload.

So: no OAuth consent screen, yes installation approval.

## Installing it in Outlook

Start ReplyBrief first: the add-in's manifest, task pane and icons are all served by it.

Sideloading follows Microsoft's documented manual flow, which is the same for Outlook on the web,
new Outlook on Windows and classic Outlook:

1. In your browser go to **<https://aka.ms/olksideload>**. Outlook on the web opens and the
   **Add-Ins for Outlook** dialog appears after a few seconds.
2. Select **My add-ins**.
3. At the bottom, in the **Custom Addins** section, select **Add a custom add-in**, then **Add from
   File**. ("Add from URL" is no longer available for custom Outlook add-ins.)
4. Choose the manifest at `<ReplyBrief output folder>\outlook\manifest.xml`, or the copy in
   `integrations\outlook\manifest.xml`, and accept the prompts.
5. Open a message in Outlook, click **ReplyBrief** on the **Message** tab, paste the pairing code
   ReplyBrief is showing, then press **Capture conversation**.

The add-in sideloaded this way shows up in supported desktop clients too. In classic Outlook on
Windows a manually sideloaded add-in can take up to 24 hours to appear, because of client-side
caching. Repeat the sideload whenever the manifest's `<Id>` or `<Version>` changes.

If **My add-ins** shows no Custom Addins section at all, that is the administrator blocking it.

### Status of this add-in: what has and has not been verified

Being precise about this, because it is easy to overstate:

* **Verified:** the manifest passes Microsoft's own schema validator
  (`npx office-addin-manifest validate`, see below), and the connector, asset serving and manifest
  structure are covered by automated tests that run against a real Kestrel instance.
* **Verified:** the task pane's JavaScript behaviour is exercised against a simulated Office runtime by
  `tests/outlook/taskpane.test.js`, run with `bun test tests/outlook`: no capture without a click,
  correct read of `item.subject` / `item.from` / `body.getAsync(Office.CoercionType.Text)`, no stale send
  on item switch, honest reporting after a send, busy and oversized handling, no leaking of Office error
  text, and message metadata rendered as text rather than markup.
* **Not verified:** nobody has sideloaded this build into a real Outlook client. Claims about new
  Outlook, Outlook on the web and classic Outlook are inferred from the declared Mailbox 1.5
  requirement and Microsoft's documented support, not from an observed run. Treat first-time
  installation as an untested step.

### Validating the manifest

```
npx office-addin-manifest validate integrations/outlook/manifest.xml
```

Without Node installed, the same validator runs through Bun:

```
bun x office-addin-manifest validate integrations/outlook/manifest.xml
```

It should end with `The manifest is valid.` This catches the class of mistake that an XML parser
cannot see, such as an element name that is not in the schema or a `resid` that resolves to nothing.
The structural guarantees are also asserted by `OutlookManifestTests`, which run with
`dotnet test` and do not need the CLI. The last validation of this file ended with:

```
Valid Manifest Schema: Your manifest does adhere to the current set of XML schema definitions for Add-in manifests.
The manifest is valid.
```

### The development certificate

The task pane is served over HTTPS because Outlook requires it. Kestrel uses the ASP.NET Core
development certificate, which is **not trusted by default**. Until it is trusted, Outlook will not
render the pane.

ReplyBrief deliberately does not change your trust store. Trust the certificate yourself, as an
explicit, reviewable action:

```
dotnet dev-certs https --trust
```

That modifies the certificate store in `CurrentUser\My` and `CurrentUser\Root` for your account only.
Remove it again with `dotnet dev-certs https --clean` if you want to undo it. In a managed
environment, ask your administrator to issue a certificate for `localhost` through your normal channel
instead.

## What the add-in does and does not do

* Captures the **selected message only**. `item.subject`, `item.from` (falling back to
  `item.sender`) and `item.body.getAsync(Office.CoercionType.Text)`. The task pane says so on screen.
* **Thread history, attachments, inline images and linked files are not included.** If the reply
  depends on earlier messages, capture them too.
* Never captures without a click. Selecting a different message while the pane is pinned only updates
  the preview; it does not send anything.
* If the reading pane switches messages **while the body is being read**, the in-flight capture is
  dropped and the pane says so, rather than sending the previous message under the new one's name.
  If the switch happens **after the request was sent**, the pane reports that the earlier message was
  sent and that the pane has moved on, because at that point it genuinely has.
* Refuses messages over 60,000 characters instead of silently truncating them.
* Never forwards a raw Office or network error message into the DOM or the console; the user sees a
  fixed, written explanation instead.
* Keeps the pairing code in memory and in a password field. It is never written to `localStorage`,
  `sessionStorage`, a cookie or a URL. **Forget pairing code** clears it.

### What this add-in deliberately does not do

* **No pinning.** The manifest declares no `TaskpaneId`, so the task pane cannot be pinned to a
  message and survives an item switch by no means of its own. The nested `VersionOverrides` 1.1 block
  is required by Mailbox 1.5, not by pinning. Pinning would need a task pane ID declared in the same
  block as the requirement set, which this add-in does not do.
* **No automatic capture**, not even when the pane is pinned.
* **No AI call from the add-in.** Analysis happens only if and when the desktop app decides to do it.

## The capture request

`POST https://localhost:47831/capture`

```http
X-ReplyBrief-Pairing: <the code ReplyBrief is showing>
Content-Type: application/json
Origin: https://localhost:47831
```

```json
{
  "source": "outlook",
  "title": "Re: roadmap",
  "participants": "Dana Example <dana@example.com>",
  "content": "Can you send the latest numbers before Thursday?",
  "contextUsed": "Outlook reading pane: the selected message's subject, sender and body text."
}
```

Responses are `202 {"status":"accepted"}` or an error with
`{"status":"rejected","error":"<safe reason>"}`:

| Status | Meaning |
| --- | --- |
| `400` | Malformed JSON, unknown source, or a capture that Core rejects |
| `401` | Missing or stale pairing code |
| `403` | Request did not come from this task pane's origin |
| `413` | Body over 300 KB |
| `415` | Not `application/json` |
| `429` | The bounded queue is full; ReplyBrief is busy |
| `503` | The host is shutting down |
| `404` | Unknown asset |

## Building and testing

```
dotnet build src\ReplyBrief.Connectors\ReplyBrief.Connectors.csproj
dotnet test  tests\ReplyBrief.Connectors.Tests\ReplyBrief.Connectors.Tests.csproj
bun test tests/outlook
```

The .NET tests start the real host on an operating system assigned loopback port and talk to it over
real HTTPS. They do not disable certificate validation and they do not touch the trust store: the test
client accepts only the certificate that is already installed in `CurrentUser\My`, presented over
loopback. This means the tests need `dotnet dev-certs https` to have been run at least once, but they
do not need it to be trusted.

`tests/outlook/taskpane.test.js` needs only Bun (`bun test tests/outlook`, run from the repository
root). It loads the real `taskpane.js` against a fake Office runtime and a minimal DOM, and covers the
behaviour the .NET tests cannot reach: nothing is sent without a click, the pairing code is never
persisted, a switch before the send drops the stale capture, a switch after the send still reports the
send truthfully, oversized and empty bodies are refused, a busy desktop is reported as busy, Office
error text never leaks, and message metadata is inserted as text rather than markup.

Validate the manifest against Microsoft's own schema (see the earlier section):

```
bun x office-addin-manifest validate integrations/outlook/manifest.xml
```

## Deployment caveats

* **One desktop app per machine.** Port `47831` is fixed because the manifest hard-codes it. A second
  ReplyBrief instance will fail to bind; the first one keeps running.
* **The certificate must be trusted on every machine** that installs the add-in, and the manifest has
  to be sideloaded on each of them. There is no code signing or distribution story here.
* **`localhost` means this machine.** The add-in cannot be used from another machine, or over a
  remote session that does not share the loopback interface.
* **Administrators can block it.** Custom add-ins are an Exchange-governed surface; ReplyBrief has no
  influence over that policy.
* **`SupportUrl` points at `https://localhost:47831/outlook/support.html`**, a small local help page
  served by the desktop app. That satisfies the manifest requirement and is reachable only while
  ReplyBrief is running, which is a deliberate trade: replace it with a real hosted support page if
  this is ever distributed beyond the machine it was built on.
* **Assets are read from disk per request** and served with `no-store`, so you can edit
  `integrations/outlook/*`, rebuild and see the change without restarting anything except the browser
  side of the pane.
* **CSP `frame-ancestors` and `frame-src` are deliberately broad**, listing the Office and Outlook
  origins rather than a single one: Office embeds a task pane from several different hosts and tenant
  domains and the runtime can open a nested frame, and a policy that is too narrow blanks the pane
  instead of degrading. ReplyBrief frames nothing itself, so this is an allowance rather than a need.
  It is not a blanket `*`, and `script-src` stays limited to the hosted office.js runtime and this
  origin.
