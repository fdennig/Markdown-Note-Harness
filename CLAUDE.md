# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A Sublime Text 4 package (lives in `Packages/francis_sublime_package/`) that extends Markdown editing with wiki-linking, todo syncing, email reply, and Office-to-Markdown conversion. It depends on the MarkdownEditing package and overrides some of its commands.

## Architecture

**Two Python plugin modules, loaded automatically by Sublime Text:**

- `francis.py` — Five `sublime_plugin` commands, plus shared frontmatter helpers (`parse_frontmatter`, `frontmatter_field`, `read_frontmatter`, `find_frontmatter_symbol_files`):
  - `OpenLinkedFileCommand` (TextCommand) — opens the URL or file path from a Markdown link `[text](path)` under the caret. Bound to F12 in markdown link scope.
  - `OpenFrontmatterAttachmentCommand` (TextCommand) — opens attachment files referenced in YAML frontmatter (prepends `./attachments/`). Bound to F12 in `entity.name.attachment.yaml` scope.
  - `ConvertToMarkdownCommand` (WindowCommand) — converts .docx/.pptx/.pdf to Markdown using Pandoc or marker-pdf. Hardcoded tool paths point to `C:\portablePrograms\`.
  - `ReplyToEmailCommand` (TextCommand) — parses `from_email` and `subject` from YAML frontmatter and opens a `mailto:` link.
  - `OpenFrontmatterSymbolCommand` (WindowCommand) — finds files by an indexed frontmatter symbol and either browses them in a quick panel with frontmatter-folded previews or opens them all in tabs. Symbol-agnostic: takes `symbol`, `label`, and `pick` args. Registered for `toreply_note` in `francis.sublime-commands`.
  - `BrowseEmailsCommand` (WindowCommand) + `EmailListInput` (ListInputHandler) — browses all emails, narrowing by `/toreply`, `/priority`, `/personal`, and swapping the searchable field with `/subject` / `/sender`. Every `/` row lives in the same list as the emails, so typing `/` just fuzzy-filters to them; there is no callback for typing, so a keystroke cannot switch modes. Choosing one returns a new handler, which is what draws the chip and re-populates the list. `preview()` is the `on_highlight` equivalent and does the transient open plus fold. `EmailListInput.depth` is tracked separately from `applied` because a field flip is a step that adds no filter.

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

## Key Design Details

- TodoSync identifies todos by their text content (the `text` field is the unique key). Duplicate detection checks both `open_todos` and `resolved_todos` lists.
- Urgency is toggled by adding/removing a `!` prefix to todo text. Changes propagate bidirectionally: editing in source files updates the DB, editing in the report updates source files.
- File paths in the database are stored as relative paths (from `UNDPnotes/`) without extension.
- The `EXCLUDE_FOLDERS` list in TodoSync.py controls which folders are skipped during todo collection.
- To make a new frontmatter flag findable: (1) add the `key: value` pair to the pinned alternation in `YAMLEmailFrontmatter.sublime-syntax`; (2) add a matching `s/^key\s*:\s*value$/name_note/;` line to `symbolIndexTransformation` in `Symbol List - Email Flags.tmPreferences`; (3) point an `open_frontmatter_symbol` entry in `francis.sublime-commands` at that token. No Python changes needed. Steps 1 and 2 must stay in step — a pair scoped but not transformed lands in the index under its raw text.
- Email discovery opens **no files**. Paths come from the symbol index and the displayed date/sender/subject are parsed out of the filename, whose format the Outlook macro fixes as `<prefix><yyyy-MM-dd>_<HHNN> @<sender> - <subject>.md`. Cost is therefore flat in vault size — reading frontmatter per file cost ~23ms each (cold), i.e. ~45s across 2000 emails. `read_entry_fields` is the fallback for names that don't match. The filename subject is `Left$(subject, 60)` of the frontmatter one, so it can be truncated; nothing else is lost.
- `ListInputHandler.list_items()` is called once per handler activation (`setup_` in `sublime_plugin.py`), after which filtering runs natively — so typing costs nothing in Python regardless of list size. Matching is against `ListInputItem.text` only, so exactly one of sender/subject is searchable at a time — hence the `/subject` flip, which rebuilds the list with the fields swapped.
- The quick panel and input handler are one-way doors: a list goes in, an index (plus an optional `Event` of modifier keys, via `WANT_EVENT`) comes back. Neither the typed filter text nor the set of rows surviving the filter is exposed — not in `sublime.py` and not in the internal `sublime_api` either. So "act on everything currently listed" cannot be implemented, and anything needing the query text must own its own input step. `show_input_panel`'s `on_change` is the only place typed text reaches Python.
- Symbol names must be globally unique: `window.symbol_locations()` matches on name alone with no scope filter, so a bare value like `high` or `email` could collide with a symbol indexed from any other file in the project. Hence the `_note` suffix scheme. Markdown headings are indexed too, but MarkdownEditing's `symbolIndexTransformation` leaves their leading `#` intact, so they cannot collide with these.
- Keybindings use Sublime's scope-based context system — F12 dispatches to different commands based on cursor scope, and Tab triggers wiki-link creation only inside `[brackets]` in markdown files.

## Development Notes

- No build system or tests — this is a Sublime Text package loaded directly from the `Packages/` directory.
- To reload after changes: save the `.py` file and Sublime Text auto-reloads the plugin. Check the Sublime console (`Ctrl+~`) for errors.
- All print/log output goes to the Sublime console. TodoSync also writes to `Data/Logs/TodoSync.log`.
- The codebase targets Python 3.8 (Sublime Text's embedded Python). Use `str.format()` instead of f-strings.
