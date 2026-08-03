# francis_sublime_package

A Sublime Text 4 package that turns a folder of Markdown notes into a working
inbox: wiki-linking, todo syncing, Office-to-Markdown conversion, and a bridge
back to the Outlook Classic messages the email notes came from.

Depends on **MarkdownEditing**, and overrides some of its commands. Lives in
`Packages/francis_sublime_package/` and is loaded directly from disk — there is
no build step.

`CLAUDE.md` is the detailed reference: architecture, the macro/plugin contract,
and the design constraints that are expensive to rediscover. This file is the
map.

## What's in here

| file | what it does |
|---|---|
| `francis.py` | Markdown link and attachment opening, Office→Markdown conversion, email browsing, and the Outlook actions |
| `TodoSync.py` | Collects unchecked todos across the vault into `todoUNDP.md` and writes resolutions back to the source files |
| `outlook_bridge.ps1` | PowerShell COM bridge to Outlook Classic. Inert to Sublime; reached only via `subprocess` from `francis.py` |
| `MarkdownExtended.sublime-syntax` | Adds a scope for unchecked todos so they reach the symbol index |
| `YAMLEmailFrontmatter.sublime-syntax` | Adds scopes for the email frontmatter fields the macro writes |
| `Symbol List - *.tmPreferences` | Rewrites those scopes into the `_note` symbol tokens the commands look up |
| `outlook_macros/` | The VBA that produces the email notes, imported by `git subtree`. Runs in Outlook, not in Sublime |
| `wiki_page_FOR_MARKDOWN_EDITING.txt` | A patched copy of MarkdownEditing's `wiki_page.py` — see below |

## Commands

All are in the command palette; the ones with a keybinding say so.

**Email notes**

- *Email: Browse* — every indexed email in one list, narrowed by typing `/toreply`,
  `/priority` or `/personal`, with `/subject` and `/sender` swapping which field
  the search matches. Highlighting a row previews it with the frontmatter folded.
- *Email: Open all to-reply* — opens every to-reply note in tabs.
- *Email: Open original in Outlook* — `F12` with the caret in the body of a note,
  anywhere that isn't a link — and the reply / reply-all / forward variants. Acts
  on the live Outlook item the
  note was exported from, so a reply lands in the original thread. Outlook writes
  the quoted block and the threading headers itself.
- *Email: Resolve in Outlook (diagnostic)* — runs the whole lookup and touches
  nothing. The safe probe when a note won't open.
- *Reply to email* — the `mailto:` fallback. Needs nothing from Outlook, but starts
  a new conversation rather than threading.

**Todos**

- *Markdown: Collect TODOs* / *Markdown: Resolve Checked TODOs*. Also run
  automatically on save, close and load of `todoUNDP.md`.

**Other**

- *Convert Office File to Markdown* — .docx / .pptx via Pandoc, .pdf via marker-pdf.
- *Open Markdown Link* — `F12` on a Markdown link, or on an attachment name in
  frontmatter. `F12` on a `[[wiki link]]`, a reference link or a footnote stays
  MarkdownEditing's. *Open original in Outlook* is on the same key but subtracts
  every one of those scopes, so each still gets its own F12.
- *Select wiki reference* — `Tab` inside `[brackets]` in a Markdown file.

`super+alt+z` opens the current repo in Sublime Merge.

## The Outlook round trip

```
Outlook Classic ──ExtractEmail*──▶ note.md ──▶ browse / reply / forward
   (COM server)                    carries      in Sublime
        ▲                          entry_id         │
        └───── outlook_bridge.ps1 ◀── subprocess ───┘
```

The macros in `outlook_macros/` write each note's frontmatter with a foldable
`outlook:` list of back-references (`entry_id`, `store_id`, `message_id`,
`conversation_id`). `OutlookMailActionCommand` reads those and asks Outlook to act
on the real item — the only way to reply *inside* a thread, since `mailto:` cannot
set `In-Reply-To` or Outlook's conversation properties.

A note carrying that block resolves in one `GetItemFromID` — `store_id` included,
without which Outlook only searches the default store and archives resolve to
nothing.

**Legacy notes** — everything exported before the macro grew the `outlook:` block
— are found by their `subject` instead, narrowed by `from_email` and then `date`.
Outlook's Inbox, Sent Items and Deleted Items are searched; a unique survivor is
opened, and anything else — no match or several — hands the subject to Outlook's
own search box and raises the window, which is a better answer than a wrong guess.

Measured over a random sample of 40 legacy notes: **39 open directly, 1 lands in
the search box.** Typically 1–4s; the worst realistic case — a subject shared by 29
messages in one thread, resolved by sender and date — took 10s.

Two details worth knowing, because both look like bugs otherwise:

- The subject is matched as a `LIKE` on its longest word and then compared exactly
  *after sanitising*, never as an equality filter. Frontmatter subjects written
  before 2026-08 had `: ? / \ " < > | [ ]` stripped by the filename sanitiser, so
  `Re: Draft PS Invitation` is stored as `Re Draft PS Invitation`.
- Note dates are occasionally an hour out from Outlook's received time. No filter
  is ever allowed to empty the candidate set, so a drifted date costs you the
  search box rather than a wrong message.

`sent:` notes reach neither path: `.Send` re-keys the item, so they have no
`entry_id`, and their date is a send time the stored copy does not share. A note
whose message has genuinely been archived or deleted reports and stops.

## `wiki_page_FOR_MARKDOWN_EDITING.txt`

Kept as `.txt` so Sublime does not load it: it is a modified copy of
MarkdownEditing's `wiki_page.py`, meant to *replace* that file inside the
MarkdownEditing package rather than to run alongside it. It makes wiki links and
back links search the whole project, and makes the make-page-reference command
double single brackets into wiki brackets and offer every Markdown file for
insertion. (The odd `self.view.insert` calls that made this work were never fully
understood.)

## Next steps

- Capture the other recipients in the macro, so reply-all has real addresses to
  work from and the note records who was on the thread. Remember to include the
  body text.
- Backfill: the subject search already knows the ids it found, so writing them into
  the note would make the next open an O(1) lookup. Not built — it edits your notes.
- A `message_id` fallback, for archived mail whose `entry_id` Outlook has re-keyed.
- Keybindings for the reply / reply-all / forward variants. *Open* has `F12`; the
  other three would need a key each, since scope cannot tell them apart.
