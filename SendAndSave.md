# Send & Save Email Macro

`SendAndSaveEmail` sends the email you are composing **and** saves a copy of it to your
Obsidian vault as a Markdown note — in one action. The note uses the same YAML + body
layout as the notes produced by `ExtractEmail`, so sent and received mail live side by side
in your vault.

## What it does

When run from an open compose window, the macro:

1. Reads the draft (subject, body, recipients, attachments) and resolves recipient names.
2. Builds a Markdown note with YAML frontmatter:
   - `class: email`
   - `sent: true` — lets you filter outgoing notes in Obsidian
   - `subject`, `date` (send time), `from` / `from_email` (you), `to:` list, `tags:`, `attachments`
3. Saves the note as UTF-8, named:
   `Sent_<yyyy-MM-dd>_<HHNN> @<first recipient> - <subject>.md`
   (outgoing mail is keyed on the recipient, since the sender is always you).
4. Sends the email.
5. Opens the new note in Sublime Text via `subl`.

The note is saved **before** the email is sent, so if the file write fails the mail does not go out.

## Installation

This macro is part of the same `SaveEmail` / `SaveUtilities` / `USER_CONFIG` modules as
`ExtractEmail` — if those are already installed, no new modules are needed. Just make sure
your three modules contain the latest versions of the `.vba` files in this repo.

### Configure the filename prefix (optional)

`USER_CONFIG` defines the prefix for sent-note filenames:

```vb
sentFileNameStartChr = "Sent_"
```

Change it if you want a different prefix; leave it as-is for the default.

## Wiring it to a shortcut

Sending happens from the **compose window** (the Inspector), which has its own toolbar
customization — separate from the reading-list (Explorer) toolbar where `ExtractEmail` lives.

1. Open a new email.
2. **File ▸ Options ▸ Quick Access Toolbar** (or the QAT dropdown ▸ **More Commands**).
3. Set **Choose commands from** → **Macros**.
4. Select `Project1.SaveEmail.SendAndSaveEmail` → **Add**.
5. Optionally **Modify** to give it an icon and a friendly name.

The button gets an **Alt + number** access key (e.g. `Alt+4`). From then on: compose your
email, press that shortcut, and the note is filed, the email sent, and the note opened in
Sublime — in one go.

## Requirements

- Outlook Classic on Windows (VBA).
- The `subl` command-line tool on your `PATH` (to auto-open the note). Without it the note
  still saves; it just won't pop open.

## Limitations

- Runs from an open compose window only (uses `ActiveInspector.CurrentItem`).
- Mail items only (`Class = 43`).
- The draft has no send timestamp yet, so the note's `date` is stamped with the current time
  at the moment you run the macro.
