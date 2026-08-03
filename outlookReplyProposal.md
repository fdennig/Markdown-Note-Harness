# Proposal: thread-aware Reply / Reply-All / Forward via Outlook Classic

> ## Status: tier 1 rebuilt and working; tiers 2–4 still a design
>
> **Rebuilt and verified against the real profile (2026-08-02):**
> `outlook_bridge.ps1` and `OutlookMailActionCommand`, doing **`entry_id` +
> `store_id` only** — one `GetItemFromID`, then `Display` / `Reply` /
> `ReplyAll` / `Forward`. Five palette entries, no keybinding. Confirmed live:
> resolve, reply, `not_found` on a re-keyed EntryID, `bad_request` on a missing
> id or a missing request file, and `unsupported_item` on an `IPM.Appointment`.
> An appointment falls back from `ReceivedTime` to `Start` correctly.
>
> **Everything below about tiers 2–4** — `message_id`, the ±2 minute window,
> `ui_search` — plus **backfill (§3c)** and the **`mailto` floor (§3d)** remains
> unbuilt design. §1d is done (the macro now writes the raw subject).
>
> **One correction to the closing paragraph:** Outlook *is* reachable from the
> Linux side of this box. `/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe`
> drives the real profile from WSL, so COM paths can be tested directly rather
> than reasoned about — the only constraint is that scripts and request files
> must sit under a Windows-visible path, since `powershell.exe` cannot read
> `/home`. That is how the above was verified.
>
> What survives from the first attempt:
>
> - **The macro side, which works.** `SaveEmail.vba` writes the `outlook:`
>   sequence, and `outlook_refs()` in `francis.py` reads it. Both verified
>   against real exported notes.
> - **`FRONTMATTER_READ_CHARS` raised to 16384**, because a `store_id` runs to
>   ~560 hex characters and `parse_frontmatter` needs the closing `---` inside
>   the read window.
> - **Everything below**, with the corrections marked *"corrected by testing"*
>   being the valuable part — each one cost a redesign.
>
> **What was proven to work**, end to end against real Outlook, before removal:
> `GetItemFromID` with a store id; a ±2 minute received-time window over an
> `Items.Restrict` folder walk; narrowing candidates by `message_id` /
> `conversation_id` / sender / subject (`narrowed 4 -> 1 by message_id`);
> `Reply()` + `Display()` producing a correctly threaded compose window. All
> four paths resolved to the identical `EntryID`.
>
> **What was never made to work:** the Sublime side. Pressing the keybinding
> produced nothing in the console at all, which was never diagnosed — the
> likely candidates being that the plugin host never reloaded `francis.py`, or
> that this package has **no `.python-version` file** and so runs under the
> Python **3.3** plugin host rather than the 3.8 assumed elsewhere in
> `CLAUDE.md`.
>
> The 3.3 half of that is now **confirmed**: both hosts ship
> (`plugin_host-3.3.exe`, `plugin_host-3.8.exe`) and of the installed packages
> only `AutomaticPackageReloader` carries a `.python-version`. The rebuild is
> written to 3.3 rather than opting this package into 3.8, since that flag
> would move `TodoSync.py` too. If a palette entry ever goes silent again, that
> is still the first thing to check.
>
> **The three environment facts that shaped the design**, all measured on the
> real profile and none of them guessable:
>
> 1. `AdvancedSearch` is unusable — MAPI's generic `"The operation failed."`
>    for every scope, filter dialect and store, while Outlook's own search box
>    works fine. `Items.Restrict` and `GetTable` are unaffected.
> 2. The sender cannot be filtered on: `urn:schemas:httpmail:fromemail` and
>    `[SenderEmailAddress]` hold the **X500 DN** for Exchange senders, so a
>    sender clause returns zero rows for internal mail.
> 3. The mailbox keeps roughly **12 months**. July 2025 held 6 messages against
>    ~7,470 for Aug–Dec 2025. Notes older than that have no message left to
>    find, so `not_found` is the expected answer, not a fault.

Follow-on to `sendEmailImplementationNotes.md`, which established the conclusion:
**don't construct a reply — ask Outlook to reply to the original item.** This
document turns that into a concrete design across the three layers that have to
agree: the VBA exporter, a PowerShell COM bridge, and the Sublime plugin.

The shape of it:

```
Outlook Classic  ──ExtractEmailCore──▶  note.md   (now carries entry_id / store_id /
    (COM server)                                   message_id / conversation_id)
        ▲                                              │
        │                                              │ F12-ish keybinding
        └──── outlook_bridge.ps1 ◀── subprocess ────  francis.py
              Reply/ReplyAll/Forward + Display()      OutlookMailActionCommand
```

Nothing about which client is *default* matters. The `Outlook.Application` ProgID
is registered by the classic Win32 client only; New Outlook does not expose it.
So `New-Object -ComObject Outlook.Application` attaches to a running classic
instance or launches one — which is exactly the requested "fire up Outlook
classic", achieved by construction rather than by detection.

---

## 1. What the macro must carry down

*(Implemented — see `outlook_macros/vba/SaveEmail.vba`.)*

Four new back-references, written at the **end** of the block (after
`attachments:`) and as a **YAML sequence** rather than as four top-level keys.
The sequence is the same shape as `attachments:`, which means it folds to a
single line: these values are long and opaque, and nobody wants them on screen
while reading the rest of the frontmatter.

```yaml
outlook:
  - entry_id: 00000000A1B2C3…
  - store_id: 0000000038A1BB…
  - message_id: "<AM0PR03MB1234ABCD@eurprd03.prod.outlook.com>"
  - conversation_id: 5A3F91C2…
```

That parses strictly as a list of single-key maps. The one cost is that
`frontmatter_field` cannot see them — it anchors keys at column 0 — so reading
them takes the small dedicated parser in §3a.

Why all four rather than just one:

| key | lookup cost | survives | used for |
|---|---|---|---|
| `entry_id` + `store_id` | O(1) `GetItemFromID` | until the item moves store/folder | the 99% path |
| `message_id` | store search, seconds | forever | after archiving/moving |
| `conversation_id` | — | forever | disambiguating multi-hit searches |

`store_id` is not optional decoration: `GetItemFromID(entryId)` without it only
looks in the default store, so anything in an archive or a shared mailbox
resolves to nothing. The two-argument form is what makes tier 1 reliable.

`conversation_id` never drives a lookup on its own (`PR_CONVERSATION_ID` is a
binary property and DASL restrictions against binary props are not dependable).
It earns its place as a *filter over candidates already in hand* — when the
legacy search returns five same-subject messages, one `-eq` comparison in
PowerShell picks the right thread.

### 1a. `SaveEmail.vba` — `ExtractEmailCore`

Insert immediately before the closing `---` (currently line 159):

```vba
        ' --- Outlook back-references: let the plugin act on the real item ---
        resultString = resultString & "outlook:" & vbCrLf
        resultString = resultString & "  - entry_id: " & GetEntryID(oMail) & vbCrLf
        resultString = resultString & "  - store_id: " & GetStoreID(oMail) & vbCrLf
        Dim internetMsgId As String
        internetMsgId = GetInternetMessageId(oMail)
        If Len(internetMsgId) > 0 Then
            resultString = resultString & "  - message_id: " & YamlQuoted(internetMsgId) & vbCrLf
        End If
        resultString = resultString & "  - conversation_id: " & GetConversationID(oMail) & vbCrLf
```

plus two helpers alongside `GetSenderSmtp`:

```vba
Private Function GetStoreID(oItem As Object) As String
    On Error GoTo Fail
    ' oItem.Parent is the containing folder; its StoreID is what the
    ' two-argument GetItemFromID needs to find items outside the default store.
    GetStoreID = oItem.Parent.StoreID
    Exit Function
Fail:
    GetStoreID = ""
End Function

Private Function GetInternetMessageId(oMail As Outlook.MailItem) As String
    On Error GoTo Fail
    ' PR_INTERNET_MESSAGE_ID. Absent on drafts and on some internal-only items.
    GetInternetMessageId = oMail.PropertyAccessor.GetProperty( _
        "http://schemas.microsoft.com/mapi/proptag/0x1035001E")
    Exit Function
Fail:
    GetInternetMessageId = ""
End Function
```

### 1b. `ExtractMeetingCore`

Same `entry_id` / `store_id` pair from `oAppt`. Meetings can't be replied to
usefully, but it makes "open the original in Outlook" work for them, which is
the third command below and comes free.

### 1c. `SendAndSaveEmail` — a genuine limitation

A draft's `EntryID` is invalidated by `.Send` (the item moves to Sent Items and
gets a new one), and `Message-ID` is assigned by the transport, so neither
exists at the moment we write the note. What *is* available, after an explicit
`oMail.Save` before the capture block:

```vba
    oMail.Save   ' assigns a ConversationID to a fresh draft
    …
    resultString = resultString & "outlook:" & vbCrLf
    resultString = resultString & "  - conversation_id: " & conversationId & vbCrLf
```

So **sent notes get `conversation_id` only** and fall through to the search
tiers. That is the right trade: replying to your own sent mail is rare, and the
common action on a sent note — forward it — still lands on the right thread via
the search fallback. Chasing the post-send EntryID would mean an `ItemAdd` event
handler on Sent Items, which is a lot of machinery for that payoff.

### 1d. A latent bug this surfaces

`ExtractEmailCore` runs `ReplaceCharsForFileName` on `temporarySubjectLineString`
*before* writing it to frontmatter (line 93 → line 128), so `subject:` is
already stripped of `: ? / \ " < > | [ ]`. "RE: Q1/Q2 budget" is stored as
"RE Q1Q2 budget". Filename sanitising is leaking into note data.

That's invisible today, but it makes any exact subject match against Outlook
impossible, which is precisely what the legacy fallback wants to do. The fix is
small and changes nothing indexed:

```vba
        ' keep the raw subject for the note, sanitise only for the filename
        Dim rawSubject As String
        rawSubject = oMail.Subject
        temporarySubjectLineString = rawSubject
        ReplaceCharsForFileName temporarySubjectLineString, ""
        …
        resultString = resultString & "subject: """ & Replace(rawSubject, """", "''") & """" & vbCrLf
```

Worth doing regardless — but note it only helps notes exported *after* the
change, so the fallback below is designed not to depend on it.

---

## 2. The bridge: `outlook_bridge.ps1`

*(Implemented. Two departures from the sketch below, both deliberate: a
`resolve` action that runs the whole lookup chain and touches nothing — safer
than `open` for repeated testing — and a settling poll rather than a
first-non-zero one, see "Resolution tiers".)*

A single file at the package root, invoked as

```
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass \
    -File outlook_bridge.ps1 -RequestPath <tempfile.json>
```

Parameters go through a **JSON temp file, not the command line**. Subjects and
display names contain quotes, `$`, backticks and semicolons; every one of those
is a quoting hazard through `cmd`→PowerShell, and some are an injection hazard.
A file sidesteps the whole class of problem, and gives a natural place to put
the response too.

**Request:**

```json
{"action": "reply|reply_all|forward|open|ui_search",
 "entry_id": "...", "store_id": "...", "message_id": "...",
 "conversation_id": "...", "subject": "...", "from_email": "...",
 "date": "2026-07-14 09:31", "window_minutes": 2,
 "timeout_seconds": 8, "max_candidates": 25}
```

**Response** — exactly one JSON line on stdout:

```json
{"ok": true, "matched_by": "entry_id|message_id|search",
 "entry_id": "...", "store_id": "...", "message_id": "...",
 "conversation_id": "..."}
```
```json
{"ok": false, "reason": "no_outlook|not_found|ambiguous|com_error",
 "message": "...", "candidates": [{"entry_id": "...", "store_id": "...",
                                   "received": "2026-07-14 09:31",
                                   "sender": "...", "subject": "..."}]}
```

Returning the resolved ids on success is what enables self-healing (§3c).

### Resolution tiers

```powershell
param([Parameter(Mandatory=$true)][string]$RequestPath)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object Text.UTF8Encoding $false

function Emit($o) { $o | ConvertTo-Json -Compress -Depth 5 }

$req = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json

try {
    # Outlook is a single-instance COM server: this attaches if it is running,
    # launches classic Outlook if it is not. New Outlook has no ProgID, so
    # there is nothing to disambiguate.
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
} catch {
    Emit @{ ok = $false; reason = 'no_outlook'; message = $_.Exception.Message }
    exit 0
}
```

**Tier 1 — direct.**

```powershell
function Resolve-ByEntryId($ns, $entryId, $storeId) {
    if ([string]::IsNullOrWhiteSpace($entryId)) { return $null }
    try {
        if ([string]::IsNullOrWhiteSpace($storeId)) { return $ns.GetItemFromID($entryId) }
        return $ns.GetItemFromID($entryId, $storeId)
    } catch { return $null }   # moved, deleted, or wrong store
}
```

**Tier 2 — `message_id`, and Tier 3 — the time window.** Both go through one
search helper.

> **Corrected by testing.** That helper was built on `AdvancedSearch`, on the
> reasoning below. `AdvancedSearch` turned out to be **entirely unusable** on
> the target profile: MAPI's generic `"The operation failed."` for every scope
> (bare name, full path, quoted, unquoted, store root, mail folder), every
> filter dialect (`LIKE`, `ci_phrasematch`, equality) and every store — while
> Outlook's own search box worked normally. It depends on the search indexer;
> something in that path is disabled.
>
> `Folder.Items.Restrict` and `Folder.GetTable` do not depend on it — the store
> provider evaluates them — and both work. So the helper walks folders and
> restricts each one. Phase 1 reaches Inbox, Sent Items and Deleted Items by
> well-known id, so a hit there costs no enumeration; phase 2 walks the rest,
> cached stores before online ones. Measured: ~2s for a full 39-folder walk,
> ~1.2s of which is enumeration, and the realistic case never gets that far.
>
> The reasoning below is kept because it is what a reader would otherwise
> re-derive, and because `AdvancedSearch` may be the right tool on a profile
> where it works at all.

```powershell
function Find-Items($ol, $filter, $timeoutSec) {
    $found = @()
    foreach ($store in $ol.Session.Stores) {
        try { $scope = "'" + $store.GetRootFolder().FolderPath + "'" } catch { continue }
        try {
            $search = $ol.AdvancedSearch($scope, $filter, $true,
                                         "fran" + [guid]::NewGuid().ToString('N'))
        } catch { continue }   # disconnected / public-folder stores throw
        $deadline = (Get-Date).AddSeconds($timeoutSec)
        while ((Get-Date) -lt $deadline -and $search.Results.Count -eq 0) {
            Start-Sleep -Milliseconds 200
        }
        foreach ($r in $search.Results) { $found += $r }
    }
    return $found
}
```

The shipped version polls until the count **settles** (unchanged across three
reads) rather than until it is merely non-zero: `Results` populates
incrementally, so a first non-zero reading can be a partial set — which for a
tier-3 search is exactly how you return one candidate and miss the better one.
A zero count is given ~2s before that store is abandoned, and the whole sweep
shares one deadline so `timeout_seconds` bounds the total, not the per-store,
cost. It stops at the first store yielding a hit that survives the filter; the
default store comes first, so the common case is one search.

Tier 2 filter — note `message-id` is *not* content-indexed, so this falls back
to a store scan and can take a few seconds. Acceptable: it only runs when the
EntryID has gone stale.

```powershell
$mid = $req.message_id -replace "'", "''"
$filter = "@SQL=""urn:schemas:mailheader:message-id"" = '$mid'"
```

Tier 3 is the **legacy path** — notes exported before any of this existed. They
carry `from_email`, `subject` and `date`, and nothing else.

> **Corrected by testing.** The original decision here was "search on sender,
> discriminate on time". The first half is impossible: for an Exchange sender,
> both `urn:schemas:httpmail:fromemail` and `[SenderEmailAddress]` hold the
> **X500 DN**, not the SMTP address the note recorded — the same trap the
> macro's `GetSenderSmtp` exists to work around. A sender clause returns zero
> rows for internal mail, which is most of it. Measured, not reasoned: a
> `fromemail =` filter for a message that provably exists returned 0 hits.

So the filter is **time and nothing else**:

> **Discriminate on time. Narrow afterwards, in process.**

`date` is minute-exact (`Format(oMail.ReceivedTime, "yyyy-MM-dd HH:mm")`), and a
±2 minute window turns out to be very nearly a unique key — across a 24-folder
mailbox it returned exactly one message. Sender and subject then narrow the
candidate set, where the sender's SMTP address can be resolved properly per item
instead of being compared inside a filter. The subject is only ever a `LIKE` on
its longest word, since it is sanitised (§1d) and truncated in the filename, so
equality would guarantee misses.

Narrowing never empties the set: a tiebreaker that eliminates every candidate is
a wrong tiebreaker, not a decisive one, and the unnarrowed set remains the best
answer available.

```powershell
$from = ($req.from_email -replace "'", "''")
$filter = "@SQL=""urn:schemas:httpmail:fromemail"" = '$from'"
if ($req.subject_token) {
    $tok = $req.subject_token -replace "'", "''"
    $filter += " AND ""urn:schemas:httpmail:subject"" LIKE '%$tok%'"
}
```

Time filtering happens in PowerShell, not in the DASL string. Date literals in
`@SQL=` restrictions are notoriously locale-sensitive; comparing
`[datetime]` objects after the fact is boring and correct:

```powershell
$target = [datetime]::ParseExact($req.date, 'yyyy-MM-dd HH:mm',
                                 [Globalization.CultureInfo]::InvariantCulture)
$win = [double]$req.window_minutes
$hits = $found | Where-Object {
    [math]::Abs(($_.ReceivedTime - $target).TotalMinutes) -le $win
}
if ($hits.Count -gt 1 -and $req.conversation_id) {
    $narrowed = $hits | Where-Object { $_.ConversationID -eq $req.conversation_id }
    if ($narrowed.Count -ge 1) { $hits = $narrowed }
}
```

One hit → act. Zero → `not_found`. More than one → return them as `candidates`
and let Sublime ask (§3b).

**Tier 4 — hand it to Outlook's own search UI.** When nothing resolves, don't
fail silently; put the query in Outlook's search box, scoped to all folders, and
raise the window. The user picks the message and hits Ctrl+R. This is the
`ui_search` action, requested explicitly by the plugin as a second call:

```powershell
$exp = $ol.ActiveExplorer()
if (-not $exp) { $exp = $ns.GetDefaultFolder(6).GetExplorer(); $exp.Display() }
$exp.Search($req.query, 3)   # olSearchScopeAllFolders
$exp.Activate()
```

### Acting on the item

```powershell
switch ($req.action) {
    'reply'     { $compose = $item.Reply() }
    'reply_all' { $compose = $item.ReplyAll() }
    'forward'   { $compose = $item.Forward() }
    'open'      { $item.Display(); $item.GetInspector.Activate(); $compose = $null }
}
if ($compose) { $compose.Display(); $compose.GetInspector.Activate() }
```

`.Display()`, never `.Send()` — that keeps us on the benign side of the Outlook
object-model guard, and the script never touches `.Recipients` on a resolved
item, which is the other thing that can trigger a prompt. Outlook writes the
quoted `From:/Sent:/To:/Subject:` block, sets `In-Reply-To`, `References`,
`PR_CONVERSATION_INDEX` and `PR_CONVERSATION_TOPIC` itself. Nothing is
reconstructed from the Markdown, so the reply is indistinguishable from a
hand-written one.

`GetInspector.Activate()` matters in practice: without it the compose window
often opens behind Sublime.

---

## 3. The plugin: `francis.py`

### 3a. One command, four actions

```python
class OutlookMailActionCommand(sublime_plugin.TextCommand):
    """Reply / reply-all / forward / open the Outlook item this note came from.

    args: {"action": "reply" | "reply_all" | "forward" | "open"}
    """
```

It mirrors the existing `open_frontmatter_symbol` shape — one command
parameterised by args rather than four near-identical classes — and reuses
`parse_frontmatter` and `frontmatter_field` unchanged for the top-level fields
(`subject`, `from_email`, `date`).

The back-references need their own reader, because they are a sequence and
`frontmatter_field`'s `^key` anchor only sees column 0:

```python
def outlook_refs(frontmatter):
    """Read the foldable 'outlook:' sequence into a dict.

    The macro writes the back-references as a YAML sequence rather than as
    top-level keys so the block folds to one line like 'attachments:' does.
    Empty values are dropped, so a caller can just test for the key: the macro
    emits 'entry_id: ' empty rather than omitting it when the lookup failed.
    """
    refs = {}
    if not frontmatter:
        return refs
    in_block = False
    for line in frontmatter.splitlines():
        if re.match(r'^outlook\s*:\s*$', line):
            in_block = True
            continue
        if in_block:
            m = re.match(r'^\s+-\s*(\w+)\s*:\s*"?([^"\r\n]*?)"?\s*$', line)
            if not m:
                break        # first non-item line ends the sequence
            value = m.group(2).strip()
            if value:
                refs[m.group(1)] = value
    return refs
```

Resolution then reads `refs.get("entry_id")`, `refs.get("message_id")` and so
on, with the tier order of §2 falling out of which keys are present. A note that
predates the macro change simply returns `{}` and goes straight to tier 3.

The subprocess runs on a `threading.Thread`, with results marshalled back
through `sublime.set_timeout`. COM start-up can take several seconds on a cold
Outlook; blocking the UI thread for that is not acceptable. Console hiding uses
the same trick as `ConvertToMarkdownCommand`, plus `CREATE_NO_WINDOW`:

```python
CREATE_NO_WINDOW = 0x08000000

def run_bridge(request, on_done):
    def work():
        fd, req_path = tempfile.mkstemp(suffix=".json")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(request, f)
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            proc = subprocess.Popen(
                [powershell_path(), "-NoProfile", "-NonInteractive",
                 "-ExecutionPolicy", "Bypass", "-File", bridge_path(),
                 "-RequestPath", req_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                startupinfo=startupinfo, creationflags=CREATE_NO_WINDOW)
            out, err = proc.communicate()
        finally:
            try: os.unlink(req_path)
            except OSError: pass
        result = parse_last_json_line(out)
        sublime.set_timeout(lambda: on_done(result, err), 0)
    threading.Thread(target=work, daemon=True).start()
```

`bridge_path()` is
`os.path.join(sublime.packages_path(), "francis_sublime_package", "outlook_bridge.ps1")`
— the package is a loose directory, so the file is on disk and directly
executable. (If it were ever zipped into a `.sublime-package`, this becomes
`sublime.load_resource` + write-to-temp; worth a one-line comment saying so.)

`powershell.exe`, not `pwsh` — PowerShell 7's COM interop with Outlook is
unreliable, and 5.1 ships with Windows.

### 3b. Disambiguation reuses the quick panel

An `ambiguous` response carries candidates; show them exactly as the email
browser shows rows (`row_text`-style `date — sender — subject`), then re-invoke
the bridge with the chosen `entry_id`/`store_id`. This is the one-way-door
pattern already documented in CLAUDE.md — a list goes in, an index comes back —
and it fits, because we own the list.

### 3c. Self-healing legacy notes

When `matched_by != "entry_id"`, the response carries the ids we just resolved.
Write them into the note's frontmatter, so the *next* action on that legacy note
is a tier-1 O(1) lookup. This is the part that makes the fallback a migration
path rather than a permanent tax.

The note is already open in the view the command ran on, so this is a plain
`TextCommand` edit — no on-disk rewriting, no BOM handling (Sublime already
detected `UTF-8 with BOM` on load and preserves it on save):

```python
class OutlookBackfillIdsCommand(sublime_plugin.TextCommand):
    """Insert or refresh entry_id/store_id/message_id before the closing '---'."""
    def run(self, edit, fields):
        ends = self.view.find_by_selector(FRONTMATTER_END_SELECTOR)
        …  # replace existing key lines in place, else insert above ends[0]
```

`FRONTMATTER_END_SELECTOR` already exists for the fold logic — reusing it means
the insertion point is defined by the syntax rather than by a second regex.

Behind a setting, defaulting on:

```json
// francis.sublime-settings
{
    "outlook_backfill_ids": true,
    "outlook_powershell": "powershell.exe",
    "outlook_search_timeout_seconds": 8,
    "outlook_match_window_minutes": 2,
    "outlook_autosave_after_backfill": true
}
```

Saving the note after backfill triggers `TodoSyncListener.on_post_save`, which
is harmless — it only acts on `todoUNDP.md` — but the setting exists for anyone
who would rather not have a command dirty their buffer.

### 3d. The floor: keep `mailto`

`reply_to_email` stays, demoted to the last resort. If the bridge returns
`no_outlook` (classic not installed, COM blocked by policy) the reply and
reply-all actions fall through to today's behaviour rather than erroring. Its
body becomes a module function so both paths share it:

```python
def open_mailto_reply(view):    # today's ReplyToEmailCommand.run body
    …

class ReplyToEmailCommand(sublime_plugin.TextCommand):
    """Deprecated: mailto reply. Prefer outlook_mail_action, which threads."""
    def run(self, edit):
        open_mailto_reply(self.view)
```

Forward has no `mailto` equivalent worth having (it would send a Markdown
rendering of the body to a new thread), so it reports the failure instead.

### 3e. Registration

`francis.sublime-commands`:

```json
{ "caption": "Email: Reply in Outlook",      "command": "outlook_mail_action", "args": { "action": "reply" } },
{ "caption": "Email: Reply all in Outlook",  "command": "outlook_mail_action", "args": { "action": "reply_all" } },
{ "caption": "Email: Forward in Outlook",    "command": "outlook_mail_action", "args": { "action": "forward" } },
{ "caption": "Email: Open original in Outlook", "command": "outlook_mail_action", "args": { "action": "open" } }
```

`Default.sublime-keymap` — F12 is spoken for by scope-based dispatch, so:

```json
{ "keys": ["ctrl+alt+r"], "command": "outlook_mail_action", "args": { "action": "reply" },
  "context": [ { "key": "selector", "operator": "equal", "operand": "text.html.markdown" } ] },
{ "keys": ["ctrl+alt+shift+r"], "command": "outlook_mail_action", "args": { "action": "reply_all" }, … },
{ "keys": ["ctrl+alt+f"], "command": "outlook_mail_action", "args": { "action": "forward" }, … }
```

`ctrl+alt+f` collides with nothing in the default keymap but is close to
Sublime's find bindings; swap if it feels wrong in use.

**No changes to `YAMLEmailFrontmatter.sublime-syntax` or
`Symbol List - Email Flags.tmPreferences`.** The new keys are ordinary YAML and
must stay out of the symbol index — `entry_id` values are globally unique hex
blobs, and per CLAUDE.md every indexed name shares one flat namespace with
`window.symbol_locations()`. Adding them would be pure noise.

---

## 4. One change needed for a reason that isn't obvious

`FRONTMATTER_READ_CHARS = 8192` bounds how much of a note `read_frontmatter`
reads, and `parse_frontmatter` returns `None` if the closing `---` falls outside
that window. Exchange `StoreID` values run to several hundred hex characters and
`EntryID` adds ~140 more; on a note with a long attachment list this eats into
the margin. Bump to `16384` — the read is lazy and only hit by the filename
fallback path, so the cost is nil.

---

## 5. Failure modes, honestly

| what happens | why | what the user sees |
|---|---|---|
| item archived by retention policy | EntryID changes on move between stores | tier 2 finds it by `message_id`, ids get backfilled |
| item permanently deleted | — | `not_found` → Outlook search UI prefilled |
| legacy note, sender sends many mails | ±2 min window still ambiguous | quick panel with candidates |
| legacy note, no `from_email` | pre-dates that field | subject `LIKE` + time window only; more often ambiguous |
| classic Outlook not installed | no COM ProgID | `no_outlook` → `mailto` fallback |
| COM guard prompt | policy/AV state | Outlook's own dialog; `.Display()` keeps this rare |
| note is a `sent:` note | no EntryID capturable pre-send | search tiers, usually resolvable by conversation |

---

## 6. Suggested order of work

1. **VBA first**, and re-export a handful of test emails. Everything downstream
   needs real `entry_id`/`store_id` values to develop against. Include §1d.
2. **`outlook_bridge.ps1` standalone**, driven by hand-written request JSON from
   a PowerShell prompt. Start with `"action": "open"` — it exercises the whole
   resolution chain but composes nothing, so it is safe to run repeatedly.
   Verify all three tiers against: a fresh note, a note whose message you then
   move to another folder, and an old legacy note.
3. **Plugin wiring**, then backfill, then the disambiguation panel.
4. **CLAUDE.md**: the "macro/plugin contract" section needs the new keys added,
   since it is precisely the kind of cross-layer coupling that section exists to
   record — and `outlook_bridge.ps1` needs a line in Development Notes noting
   that a `.ps1` at package root is inert to Sublime and is only ever reached
   through `subprocess`.

I have not been able to execute any of the COM paths here — this is a Linux
environment and Outlook is not reachable from it — so every PowerShell fragment
above is written from the object model's documented behaviour and needs to be
run against real Outlook before it is trusted. Step 2 exists to make that
verification cheap and non-destructive.
