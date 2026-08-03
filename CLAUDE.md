# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A Sublime Text 4 package (lives in `Packages/francis_sublime_package/`) that extends Markdown editing with wiki-linking, todo syncing, email reply, and Office-to-Markdown conversion. It depends on the MarkdownEditing package and overrides some of its commands.

## Architecture

**Two Python plugin modules, loaded automatically by Sublime Text:**

- `francis.py` — Seven `sublime_plugin` commands, plus shared frontmatter helpers (`parse_frontmatter`, `frontmatter_field`, `read_frontmatter`, `outlook_refs`):
  - `OpenLinkedFileCommand` (TextCommand) — opens the URL or file path from a Markdown link `[text](path)` under the caret. Bound to F12 in markdown link scope.
  - `OpenFrontmatterAttachmentCommand` (TextCommand) — opens attachment files referenced in YAML frontmatter (prepends `./attachments/`). Bound to F12 in `entity.name.attachment.yaml` scope.
  - `ConvertToMarkdownCommand` (WindowCommand) — converts .docx/.pptx/.pdf to Markdown using Pandoc or marker-pdf. Hardcoded tool paths point to `C:\portablePrograms\`.
  - `ReplyToEmailCommand` (TextCommand) — parses `from_email` and `subject` from YAML frontmatter and opens a `mailto:` link.
  - `OpenFrontmatterSymbolCommand` (WindowCommand) — finds files by an indexed frontmatter symbol and either browses them in a quick panel with frontmatter-folded previews or opens them all in tabs. Symbol-agnostic: takes `symbol`, `label`, and `pick` args. Registered for `toreply_note` in `francis.sublime-commands`.
  - `BrowseEmailsCommand` (WindowCommand) + `EmailListInput` (ListInputHandler) — browses all emails, narrowing by `/toreply`, `/priority`, `/personal`, and swapping the searchable field with `/subject` / `/sender`. Every `/` row lives in the same list as the emails, so typing `/` just fuzzy-filters to them; there is no callback for typing, so a keystroke cannot switch modes. Choosing one returns a new handler, which is what draws the chip and re-populates the list. `preview()` is the `on_highlight` equivalent and does the transient open plus fold. `EmailListInput.depth` is tracked separately from `applied` because a field flip is a step that adds no filter.
  - `OutlookMailActionCommand` (TextCommand) — opens / replies / replies-all / forwards the live Outlook item a note was exported from, by shelling out to `outlook_bridge.ps1`. One command parameterised by `action`, the same shape as `open_frontmatter_symbol`; `action: "resolve"` touches nothing and is the safe probe. Palette only, no keybinding. Resolution is `entry_id` + `store_id` and nothing else — one `GetItemFromID`. A note with no `entry_id` (exported before the macro wrote the `outlook:` block, or a `sent:` note) and an item Outlook has since re-keyed by moving it between stores both report and stop; the search tiers that would cover them are not built.

- `TodoSync.py` — Todo collection and resolution system:
  - Scans markdown files for unchecked todo items (via Sublime's indexed symbols) and stores them in a JSON database (`todo.yml`, despite the name it's JSON).
  - Generates a summary report (`todoUNDP.md`) with Urgent, Inbox, and By File sections plus a Completed history.
  - On save of `todoUNDP.md`: processes checked items as resolutions, updates source files (checks boxes), manages urgency (`!` prefix), and handles inbox-only todos.
  - Expects a project folder named `UNDPnotes` containing `todo.yml` and `todoUNDP.md`.
  - `TodoSyncListener` (EventListener) triggers sync on save, close, and load events.

**Syntax extensions:**

- `MarkdownExtended.sublime-syntax` — extends the built-in Markdown syntax to add scopes for unchecked todo items (`entity.name.tag.todo.unchecked.markdown.line`), enabling symbol indexing.
- `YAMLEmailFrontmatter.sublime-syntax` — extends YAML syntax with scopes for email fields (sender, from_email, subject, attachments, toreply).

**Symbol list / indexing:**

- `Symbol List - ToDos.tmPreferences` — makes unchecked todos appear in Sublime's symbol index (used by TodoSync to discover todos via `view.indexed_symbol_regions()`).
- `Symbol List - Email Flags.tmPreferences` — indexes email frontmatter flags. The syntax scopes whole `key: value` pairs as `entity.name.email.flag.yaml`; this file's `symbolIndexTransformation` rewrites each pair into its own token: `toreply_note`, `personal_note`, `sent_note`, `priority_note`, `email_note`, `meeting_note`. Those tokens — not the raw keys or values — are what `window.symbol_locations()` looks up.

**Outlook macros (`outlook_macros/`):**

VBA that runs inside Outlook, not inside Sublime. It is the *producer* of every email note the plugin consumes, so the two halves live together — but nothing here is loaded or executed by Sublime. Editing a `.vba` file changes nothing until it is pasted into Outlook's VBE (Alt+F11) and the module is re-imported.

- `vba/SaveEmail.vba` — `ExtractEmail`, `ExtractEmailToReply`, `ExtractEmailWithAttachments`, `ExtractEmailPersonal`, `ExtractEmailHighPriority` all funnel into `ExtractEmailCore(toReply, saveAttachments, isPersonal, priority)`; `ExtractMeetingCore` handles appointments and `SendAndSaveEmail` sends outgoing mail and archives it with `sent: true`.
- `vba/SaveUtilities.vba` — filename sanitising (`ReplaceCharsForFileName`), `formatName` (wraps people as `[[@Name]]`), and `SaveAsUTF8`, which writes a **BOM** — hence `read_frontmatter` opening files as `utf-8-sig`.
- `vba/USER_CONFIG.vba` — vault paths and filename prefixes. Contains machine-specific absolute paths.

Imported with `git subtree` from `https://github.com/fdennig/outlook2markdown.git` (itself a fork of machado-t/Outlook2Obsidian, MIT — `outlook_macros/LICENSE` must stay with the code). To sync either way:

```
git subtree pull --prefix=outlook_macros https://github.com/fdennig/outlook2markdown.git main
git subtree push --prefix=outlook_macros https://github.com/fdennig/outlook2markdown.git main
```

**The macro/plugin contract** — change one side and the other breaks silently:

- Frontmatter values are pinned in `YAMLEmailFrontmatter.sublime-syntax`, so the macro's exact spellings matter: `toreply: true`, `personal: true`, `sent: true`, `priority: high`, `class: email`, `class: meeting`. `ExtractEmailCore` writes `toreply: `, `personal: ` and `priority: ` *empty* when the flag is off — those lines deliberately fail to match and never reach the symbol index.
- Filenames are `<prefix><yyyy-MM-dd>_<HHNN> @<sender> - <subject>.md`, meetings omitting ` @<sender>`. `EMAIL_FILENAME_RE` in `francis.py` parses this and is the reason email discovery opens no files. Changing the prefix is safe (the regex anchors on the date token); changing the separators is not.
- `read_entry_fields` reads `from`/`sender`, `subject` and `date` from frontmatter as the fallback for names that don't match.
- **Outlook back-references.** `ExtractEmailCore` closes the frontmatter with an `outlook:` sequence holding `entry_id`, `store_id`, `message_id` (omitted when unavailable) and `conversation_id`; `ExtractMeetingCore` writes the first two; `SendAndSaveEmail` can only write `conversation_id`, because `.Send` re-keys the item's EntryID and the Message-ID is stamped by the transport. They exist so that something can act on the live Outlook item inside its original thread, which no `mailto:` can do. `outlook_refs()` in `francis.py` reads them and `OutlookMailActionCommand` acts on them via `outlook_bridge.ps1`; only `entry_id` + `store_id` are used so far, so `message_id` and `conversation_id` are carried but not yet consumed. `outlookReplyProposal.md` holds the design for the search tiers that would consume them and, more usefully, what testing against a real profile disproved. Two consequences of the sequence shape: it folds to one line like `attachments:` does, which is why it is a sequence and why it goes last; and `frontmatter_field` **cannot read it** (that regex anchors keys at column 0), hence `outlook_refs`. They are deliberately **not** scoped or indexed: they are opaque per-item values and every indexed symbol shares one flat namespace.
- The note's `subject` is the **raw** Outlook subject (via `YamlQuoted` in `SaveUtilities.vba`); only the filename copy goes through `ReplaceCharsForFileName`. Before that split, filename sanitising silently stripped `: ? / \ " < > | [ ]` out of the frontmatter too. Notes exported before the change still carry the stripped form, so nothing may rely on `subject` matching Outlook exactly.

## Key Design Details

- TodoSync identifies todos by their text content (the `text` field is the unique key). Duplicate detection checks both `open_todos` and `resolved_todos` lists.
- Urgency is toggled by adding/removing a `!` prefix to todo text. Changes propagate bidirectionally: editing in source files updates the DB, editing in the report updates source files.
- File paths in the database are stored as relative paths (from `UNDPnotes/`) without extension.
- The `EXCLUDE_FOLDERS` list in TodoSync.py controls which folders are skipped during todo collection.
- To make a new frontmatter flag findable: (1) add the `key: value` pair to the pinned alternation in `YAMLEmailFrontmatter.sublime-syntax`; (2) add a matching `s/^key\s*:\s*value$/name_note/;` line to `symbolIndexTransformation` in `Symbol List - Email Flags.tmPreferences`; (3) point an `open_frontmatter_symbol` entry in `francis.sublime-commands` at that token. No Python changes needed. Steps 1 and 2 must stay in step — a pair scoped but not transformed lands in the index under its raw text.
- Email discovery opens **no files**. Paths come from the symbol index and the displayed date/sender/subject are parsed out of the filename, whose format the Outlook macro fixes as `<prefix><yyyy-MM-dd>_<HHNN> @<sender> - <subject>.md`. Cost is therefore flat in vault size — reading frontmatter per file cost ~23ms each (cold), i.e. ~45s across 2000 emails. `read_entry_fields` is the fallback for names that don't match. The filename subject is `Left$(…, 60)` of the *sanitised* subject, so relative to the frontmatter one it is both truncated and missing any `: ? / \ " < > | [ ]`; nothing else is lost.
- `ListInputHandler.list_items()` is called once per handler activation (`setup_` in `sublime_plugin.py`), after which filtering runs natively — so typing costs nothing in Python regardless of list size. Matching is against `ListInputItem.text` only, so exactly one of sender/subject is searchable at a time — hence the `/subject` flip, which rebuilds the list with the fields swapped.
- The quick panel and input handler are one-way doors: a list goes in, an index (plus an optional `Event` of modifier keys, via `WANT_EVENT`) comes back. Neither the typed filter text nor the set of rows surviving the filter is exposed — not in `sublime.py` and not in the internal `sublime_api` either. So "act on everything currently listed" cannot be implemented, and anything needing the query text must own its own input step. `show_input_panel`'s `on_change` is the only place typed text reaches Python.
- Symbol names must be globally unique: `window.symbol_locations()` matches on name alone with no scope filter, so a bare value like `high` or `email` could collide with a symbol indexed from any other file in the project. Hence the `_note` suffix scheme. Markdown headings are indexed too, but MarkdownEditing's `symbolIndexTransformation` leaves their leading `#` intact, so they cannot collide with these.
- Keybindings use Sublime's scope-based context system — F12 dispatches to different commands based on cursor scope, and Tab triggers wiki-link creation only inside `[brackets]` in markdown files.

## Development Notes

- No build system or tests — this is a Sublime Text package loaded directly from the `Packages/` directory.
- Sublime only auto-loads plugins from a package's **root**; `.py` in a subdirectory is importable but never loaded (MarkdownEditing keeps a root `plugin.py` that imports `plugins/`). `.sublime-keymap`, `.sublime-commands`, `.sublime-settings` and `.sublime-menu` are likewise read from the root only. But `.sublime-syntax`, `.tmLanguage` and `.tmPreferences` are found **recursively** — so never add one under `outlook_macros/`, since a stray symbol-list file would inject names into the index that `window.symbol_locations()` could confuse with the `_note` tokens.
- To reload after changes: save the `.py` file and Sublime Text auto-reloads the plugin. Check the Sublime console (`Ctrl+~`) for errors.
- All print/log output goes to the Sublime console. TodoSync also writes to `Data/Logs/TodoSync.log`.
- The codebase runs on the **3.3 plugin host**, not 3.8. Opting into 3.8 needs a `.python-version` file at package root and this package has none (of the installed packages, only AutomaticPackageReloader does). So: no f-strings — use `str.format()` — and no 3.4+ standard library. `subprocess.communicate(timeout=…)` and `threading.Thread(daemon=…)` both landed in 3.3 and are safe.
- `outlook_bridge.ps1` sits at package root but is inert to Sublime: a `.ps1` is loaded by nothing. It is reached only through `subprocess` from `OutlookMailActionCommand`, which finds it via `sublime.packages_path()` — that works because the package is a loose directory. Zipped into a `.sublime-package` it would have to become `sublime.load_resource()` plus a write to temp.
- Outlook COM is testable directly from a WSL shell: `/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe` drives the real profile. Scripts and any request file must live under a Windows-visible path (`/mnt/c/…`) — `powershell.exe` cannot read `/home`.
