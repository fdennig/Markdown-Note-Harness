Short answer: **none at all.** Sublime's embedded Python is a stock CPython 3.8 with a trimmed stdlib, and the `sublime`/`sublime_plugin` modules are purely editor API — buffers, views, windows, settings. There's no mail, MAPI, or COM binding anywhere in it. `pywin32` isn't bundled, and installing it into Sublime's interpreter is awkward (compiled extensions plus DLL loading).

But that doesn't matter much, because your plugin already demonstrates the real bridge: `ConvertToMarkdownCommand` shells out with `subprocess.Popen`. Sublime is a native Windows process, so anything you can drive from PowerShell you can drive from a Sublime command.

## Why mailto can never do this

Outlook's conversation grouping isn't derived from the subject line. It reads the `In-Reply-To` and `References` headers, plus Outlook's own MAPI properties `PR_CONVERSATION_INDEX` and `PR_CONVERSATION_TOPIC`. RFC 6068 restricts which headers a `mailto:` URL may set, and Outlook honors essentially subject/body/cc/bcc. There is no mailto formulation that threads correctly — you'd always get a new root conversation.

## The approach that works

Don't *construct* a reply. Ask Outlook to reply to the original item — then Outlook itself writes the quoted original block (`From: / Sent: / To: / Subject:`) and sets the conversation properties correctly. That's the whole trick.

Via COM from PowerShell, which needs nothing installed:

```powershell
$ol   = New-Object -ComObject Outlook.Application
$ns   = $ol.GetNamespace("MAPI")
$item = $ns.GetItemFromID($entryId)
$reply = $item.Reply()        # or .ReplyAll()
$reply.Display()              # opens the compose window; don't .Send()
```

Your command becomes: parse frontmatter → `subprocess.Popen(["powershell", "-NoProfile", "-Command", ...])` with the same `startupinfo` console-hiding trick you already use for Pandoc.

## What you need in the frontmatter

You need a handle on the original message, so the downloader has to capture one at export time:

- **`entry_id`** — Outlook's `EntryID`. Direct lookup via `GetItemFromID`, fastest. Downside: it changes if the item moves between folders or stores, so archived mail can break.
- **`message_id`** — the RFC 5322 `Message-ID` header. Stable forever. Look it up with a DASL restriction on `urn:schemas:mailheader:message-id`, or `AdvancedSearch`. Slower but survives archiving.

Storing both and preferring `entry_id` with a `message_id` fallback is the robust combination. Falling back to a subject+sender+received-time search when neither is present works but is genuinely brittle — same-subject threads collide.

## Two caveats that decide whether any of this is viable

**"New Outlook" has no COM interface.** The classic Win32 desktop client exposes the Outlook Object Model; the rewritten client (and Outlook on the web) does not. If you're on new Outlook, the equivalent is Microsoft Graph — `POST /me/messages/{id}/createReply`, which returns a properly threaded draft. That's a genuinely bigger lift: app registration, OAuth device-code flow, token caching. Worth confirming which client you're on before investing in the COM path.

**The Outlook security model** can prompt on programmatic access to address data. Using `.Display()` rather than `.Send()` keeps you out of the worst of it, and on a domain-joined machine with current AV registered in Security Center the guard is usually suppressed entirely.

One nice side effect of the COM route: since Outlook generates the quoted body, you don't need to reconstruct anything from your downloaded Markdown — the reply looks exactly like a hand-written one to every recipient's client.
