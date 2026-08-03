# Why replying from Sublime has to go through Outlook

Background for `outlookReplyProposal.md`, which builds on the conclusion here.
Kept because it is the *why*: the proposal assumes all of this and does not
re-argue it.

## Sublime has no mail API, and does not need one

The `sublime` / `sublime_plugin` modules are purely editor API — buffers, views,
windows, settings. No mail, no MAPI, no COM. `pywin32` is not bundled, and
installing it into Sublime's interpreter is awkward: compiled extensions plus DLL
loading, against a host whose version you do not control. (This package runs on
the **3.3** plugin host, not 3.8 — see CLAUDE.md.)

None of that matters, because `ConvertToMarkdownCommand` already demonstrates the
real bridge: `subprocess.Popen`. Sublime is a native Windows process, so anything
drivable from PowerShell is drivable from a Sublime command.

## `mailto:` can never thread

Outlook's conversation grouping is not derived from the subject line. It reads
`In-Reply-To` and `References`, plus Outlook's own MAPI properties
`PR_CONVERSATION_INDEX` and `PR_CONVERSATION_TOPIC`. RFC 6068 restricts which
headers a `mailto:` URL may set, and Outlook honours essentially
subject / body / cc / bcc.

So there is no `mailto:` formulation that threads. It always starts a new root
conversation. `ReplyToEmailCommand` still exists for notes that carry no Outlook
handle, but that is what it costs.

## Don't construct a reply — ask Outlook to reply

The whole trick. Hand the original item back to Outlook and let it write the
quoted `From: / Sent: / To: / Subject:` block and set the conversation properties
itself:

```powershell
$ol    = New-Object -ComObject Outlook.Application
$ns    = $ol.GetNamespace("MAPI")
$item  = $ns.GetItemFromID($entryId, $storeId)
$reply = $item.Reply()        # or .ReplyAll() / .Forward()
$reply.Display()              # never .Send()
```

Nothing is reconstructed from the downloaded Markdown, so the reply is
indistinguishable from a hand-written one in every recipient's client.

This is what `outlook_bridge.ps1` does.

## What the note has to carry

A handle on the original message, captured at export time by the macro:

- **`entry_id`** — Outlook's `EntryID`. `GetItemFromID`, O(1). Re-keyed when the
  item moves between stores, so archived mail stops resolving.
- **`store_id`** — required alongside it. The one-argument `GetItemFromID` searches
  only the default store, so archives and shared mailboxes resolve to nothing
  without it.
- **`message_id`** — the RFC 5322 header. Stable forever, found by a store search,
  seconds rather than microseconds. The fallback for after an archive.
- **`conversation_id`** — never a lookup key on its own; it disambiguates
  candidates a search already returned.

The macro writes all four. The plugin currently uses the first two.

## Two caveats that decided the approach

**New Outlook has no COM interface.** The classic Win32 client exposes the Outlook
Object Model; the rewritten client and Outlook on the web do not. This turned out
to be an advantage rather than a risk: because only classic registers the
`Outlook.Application` ProgID, `New-Object -ComObject Outlook.Application` attaches
to classic or launches it, and which client is *default* never has to be detected.
On a machine with only new Outlook, the equivalent would be Microsoft Graph
(`POST /me/messages/{id}/createReply`) — app registration, OAuth device-code flow,
token caching. A much bigger lift.

**Outlook's object-model guard** can prompt on programmatic access to address data.
Using `.Display()` rather than `.Send()`, and never touching `.Recipients` on a
resolved item, stays clear of it. No prompt has been seen in testing on this
profile.
