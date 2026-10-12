# Product plan

## Goal

One Windows companion for the email or chat the user is currently reading. It supplies a summary,
attention items, suggested reply points, and missing information. It never sends the reply. Inference
runs through the user's Ollama server on their network, not a cloud AI provider.

The desktop shell is WPF on .NET 8 for a low-overhead first implementation.

## Shared workflow

1. Select an email or conversation in its application.
2. Invoke ReplyBrief in that service to hand off the selected content and its coverage metadata.
3. Preview exactly what was captured in the desktop app.
4. Explicitly analyze it through the configured Ollama server.
5. Review and use the reply points as a reference while writing the actual response.

OAuth authorizes data access; it does not tell a desktop application what is open in another app.
Each connector must provide current-item context as well as access to that item's content.

## Phase 1: Desktop and Outlook

Implemented: desktop preview/manual fallback, LAN endpoint/model selection, connection check,
bounded structured analysis, consent reset on capture changes, cancellation, and safe failure states.
The Outlook add-in requests `ReadItem` and captures one selected message over a paired local HTTPS
connection. The current manifest is unpinned; pinning and automatic selection-following are not shipped.

Acceptance still requires:

- Sideload in New Outlook and Outlook on the web with a real Microsoft 365 or Outlook.com account.
- Verify Office runtime loading, local HTTPS trust, and add-in installation policy.
- Test analysis through the user's actual LAN server and chosen model with realistic email lengths.
- Confirm output remains advisory and no source change can leave stale results beside new content.
- Test two desktops against one Ollama server, unavailable server, model load delays, and cancellation.

No Microsoft Graph OAuth is needed just to read the current Outlook item. Full thread retrieval is
separate work: obtain delegated `Mail.Read` authorization when necessary, resolve current item identity,
retrieve relevant messages with explicit bounds, and label included/omitted context. Do not read an
entire mailbox merely to find the open message. New Outlook support for non-Microsoft accounts and
protected/shared messages is not guaranteed by the current add-in and needs a documented support matrix.

## Phase 2: Gmail

Google's official Workspace add-on can identify the open message and read its thread using narrow
current-message scopes. However, that add-on runs in Google's infrastructure and cannot directly call
a Windows loopback listener or private LAN Ollama host. Do not ship a cloud relay of raw message text
without an explicit product/privacy decision.

Preferred next prototype for this desktop-first product: an Edge/Chrome extension with user-triggered
current-tab access and an authenticated native-messaging bridge to ReplyBrief. Keep capture local until
analysis. A browser extension can inspect the visible conversation, but Gmail's DOM is not a stable API:
show a coverage preview and never claim that collapsed or unloaded messages were included.

Investigate OAuth-backed Gmail API retrieval for complete threads once a supported current-thread
identity handoff is established. Broad `gmail.readonly` is a restricted scope and may require Google
verification/security assessment depending on distribution and handling. Do not add it by default if
current-message access suffices.

Acceptance requires current message/thread identity tests, collapsed-message handling, stale-tab
protection, explicit permissions, authenticated browser-to-desktop delivery, and no AI requests from
the extension. Gmail is currently **not implemented**.

## Phase 3: Teams

Start with work/school accounts. A Teams message-action extension supplies the selected message;
additional chat history needs a supported identity handoff and authorized Microsoft Graph retrieval.
Teams invokes bot-based extensions through Microsoft services, so a public HTTPS endpoint may be
required. A cloud-to-local relay must not be assumed equivalent to the Outlook localhost design.

Prototype a minimal message/link identity handoff to the desktop, then let the desktop obtain tokens
through Microsoft-supported authentication and fetch bounded chat context. Keep tokens in the Windows
credential store, never ordinary settings files. Investigate resource-specific consent versus delegated
`Chat.Read`, and verify tenant consent policy before selecting permissions. Channel threads need their
own permission and context treatment; private chats and channels are not interchangeable.

Acceptance requires real Teams desktop/web testing, organization approval where applicable,
message-vs-conversation coverage labels, membership/permission failure handling, bounded pagination,
and no silent chat-history harvesting. Personal Teams support is **not assumed**. Teams is currently
**not implemented**.

## Phase 4: Distribution

- Replace developer certificate setup with an explicit, maintainable local transport/install strategy.
- Add an installer, signing, single-instance behavior, and service-specific onboarding.
- Save only non-sensitive server/model preferences; keep raw captures ephemeral by default.
- Add accessible copy-reply-points and hotkey workflows without granting message-send permissions.
- Validate Microsoft organization deployment and Chrome/Edge extension distribution requirements.
- Document supported clients, account types, context limits, and data handling at every boundary.

## Primary references

- [Outlook add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-add-ins-overview)
- [New Outlook add-in support](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook)
- [Sideload Outlook add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)
- [Gmail contextual message UI](https://developers.google.com/workspace/add-ons/gmail/extending-message-ui)
- [Gmail OAuth scopes](https://developers.google.com/workspace/gmail/api/auth/scopes)
- [Chrome native messaging](https://developer.chrome.com/docs/extensions/develop/concepts/native-messaging)
- [Teams message extensions](https://learn.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions)
- [Graph chat messages](https://learn.microsoft.com/en-us/graph/api/chat-list-messages)
- [Ollama structured outputs](https://docs.ollama.com/capabilities/structured-outputs)
- [Ollama server configuration](https://docs.ollama.com/faq)
