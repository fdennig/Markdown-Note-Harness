<#
    outlook_bridge.ps1 -- resolve a note's Outlook back-references to the live
    item in Outlook Classic and act on it.

    Inert to Sublime: a .ps1 at package root is loaded by nothing. It is only
    ever reached through subprocess, from OutlookMailActionCommand in
    francis.py. Invoked as:

        powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass `
            -File outlook_bridge.ps1 -RequestPath <tempfile.json>

    Request (JSON file, not command line -- subjects and display names carry
    quotes, '$' and backticks, every one of which is a quoting hazard through
    cmd -> PowerShell):

        {"action": "resolve|open|reply|reply_all|forward|save_attachments",
         "entry_id": "...", "store_id": "...",
         "subject": "...", "from_email": "...", "date": "2026-07-15 16:00"}

    save_attachments also takes "dir" (an existing folder) and "prefix", and
    writes every attachment there as <prefix><FileName> -- what the macro's
    ExtractEmailWithAttachments does, for EmailBatch.py. It answers with
    "saved" and "failed" name lists. It needs an entry_id: saving the files of
    a message found by guessing at its subject is not something to do unseen.

    Response: exactly one compact JSON object on stdout, always with exit 0.

        {"ok": true,  "action": "...", "matched_by": "entry_id|subject",
         "entry_id": "...", "store_id": "...", "message_class": "IPM.Note",
         "subject": "...", "received": "...", "conversation_id": "...",
         "message_id": "..."}
        {"ok": true,  "action": "...", "matched_by": "search_ui",
         "query": "subject:\"...\"", "hits": 4}
        {"ok": false, "reason": "bad_request|no_outlook|not_found|
                                 unsupported_item|com_error", "message": "..."}

    Two ways in, and the whole chain runs in this one invocation -- splitting it
    across two calls would mean two COM attaches for no gain:

      entry_id present  GetItemFromID, O(1). A failure here is reported, not
                        retried: the note carries store_id too, so it means the
                        message genuinely moved or was deleted.
      subject only      For notes exported before the macro wrote the 'outlook:'
                        block. Restrict Inbox / Sent Items / Deleted Items, then
                        narrow by sender and by received time. A unique survivor
                        is acted on; anything else -- none or several -- hands
                        the subject to Outlook's own search box and raises the
                        window, which is a better answer than a wrong guess.

    Nothing here calls .Send(). Composing goes through Reply/ReplyAll/Forward
    followed by .Display(), which keeps the script on the benign side of the
    Outlook object-model guard; .Recipients is never touched, which is the other
    thing that can raise a security prompt.
#>
param([Parameter(Mandatory = $true)][string]$RequestPath)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false

# Every exit from this script goes through Emit, so the caller can rely on
# there being exactly one JSON line to parse. 'exit' inside a function ends the
# script, which is what makes the early-return style below work.
function Emit($payload) {
    $payload | ConvertTo-Json -Compress -Depth 4
    exit 0
}

# Safety net for anything the explicit catches miss -- with $ErrorActionPreference
# set to Stop, an uncaught error would otherwise end the script silently and the
# caller would see an empty stdout it could say nothing useful about.
trap { Emit @{ ok = $false; reason = 'com_error'; message = $_.Exception.Message } }

# COM properties are absent or throw depending on item class (an AppointmentItem
# has no ReceivedTime, a draft has no Message-ID), so every read is optional.
#
# It has to be a function. PowerShell 5.1 has no try/catch *expression*, so an
# inline `(try { $m.Subject } catch { '' })` parses as a call to a command named
# 'try' -- and since these reads all sit inside other catches, it fails silently.
function Read-Prop([scriptblock]$block, $fallback = $null) {
    try {
        $value = & $block
        if ($null -eq $value) { $fallback } else { $value }
    } catch { $fallback }
}

# --- Finding a message by subject -------------------------------------------

# Frontmatter subjects written before 2026-08 went through the filename
# sanitiser, so 'RE: Q1/Q2 budget' was stored as 'RE Q1Q2 budget'. Comparing a
# live subject against a note's means putting the live one through the same
# mangling -- and it is exactly those notes this path exists for.
function ConvertTo-SanitisedSubject([string]$subject) {
    return ($subject -replace '[:?/\\"<>|\[\]]', '').Trim()
}

# The store-side filter is a LIKE on one long word, never an equality on the
# whole subject: the note's copy is sanitised, so equality would miss. Exactness
# is recovered afterwards by the sanitised comparison above.
function Get-SubjectToken([string]$subject) {
    $words = $subject -split '[^\p{L}\p{Nd}]+' | Where-Object { $_.Length -gt 2 }
    if (-not $words) { return $null }
    return ($words | Sort-Object -Property Length -Descending)[0]
}

# Readable once an item is in hand, though impossible to filter on:
# SenderEmailAddress holds the X500 DN for an Exchange sender, and
# PR_SMTP_ADDRESS belongs to the recipient table rather than the sender, so both
# come back useless and a tiebreaker built on them silently never fires.
function Get-SenderSmtp($mail) {
    foreach ($tag in '0x5D01001F', '0x5D02001F') {   # PR_SENDER_SMTP_ADDRESS, PR_SENT_REPRESENTING_SMTP_ADDRESS
        $addr = Read-Prop {
            $mail.PropertyAccessor.GetProperty(
                'http://schemas.microsoft.com/mapi/proptag/' + $tag)
        } ''
        if ($addr -like '*@*') { return $addr }
    }
    $addr = Read-Prop { $mail.SenderEmailAddress } ''
    if ($addr -like '*@*') { return $addr }
    return Read-Prop { $mail.Sender.GetExchangeUser().PrimarySmtpAddress } ''
}

# Inbox, Sent Items and Deleted Items, reached by well-known id so nothing is
# enumerated -- ~1-2s in total. Anything filed elsewhere falls through to
# Outlook's own search, which covers every folder and the online archive.
function Find-BySubject($ns, [string]$subject) {
    $token = Get-SubjectToken $subject
    if (-not $token) { return @() }

    $q = [char]39
    $filter = '@SQL="urn:schemas:httpmail:subject" LIKE ' + $q + '%' +
              ($token -replace $q, ($q + $q)) + '%' + $q
    $want = ConvertTo-SanitisedSubject $subject

    $hits = @()
    foreach ($wellKnown in 6, 5, 3) {    # olFolderInbox, olFolderSentMail, olFolderDeletedItems
        $folder = Read-Prop { $ns.GetDefaultFolder($wellKnown) }
        if (-not $folder) { continue }
        try {
            foreach ($mail in $folder.Items.Restrict($filter)) {
                $live = Read-Prop { [string]$mail.Subject } ''
                if ((ConvertTo-SanitisedSubject $live) -eq $want) { $hits += $mail }
            }
        } catch { }    # non-mail folders and disconnected stores throw; skip them
    }
    return @($hits)
}

# Sender first, then received time. Neither is ever allowed to empty the set: a
# tiebreaker that eliminates every candidate is a wrong tiebreaker, not a
# decisive one. That is also what makes the timezone drift harmless -- a note
# dated 19:20 whose message was received at 18:20 keeps its candidates and
# reaches the search box, rather than being discarded.
function Select-Unique($candidates, $req) {
    # Where-Object, not a bare @(): a PowerShell function returning an empty
    # array yields $null, and @($null) is an array of one $null -- which counts
    # as a unique match and resolves to nothing at all.
    $candidates = @($candidates | Where-Object { $null -ne $_ })
    if ($candidates.Count -le 1) { return $candidates }

    if (-not [string]::IsNullOrWhiteSpace($req.from_email)) {
        $bySender = @($candidates | Where-Object {
            (Get-SenderSmtp $_) -eq [string]$req.from_email
        })
        if ($bySender.Count -ge 1) { $candidates = $bySender }
    }
    if (@($candidates).Count -le 1) { return @($candidates) }

    if (-not [string]::IsNullOrWhiteSpace($req.date)) {
        $target = Read-Prop {
            [datetime]::ParseExact([string]$req.date, 'yyyy-MM-dd HH:mm',
                                   [Globalization.CultureInfo]::InvariantCulture)
        }
        if ($target) {
            $byTime = @($candidates | Where-Object {
                $received = Read-Prop { $_.ReceivedTime }
                $received -and [math]::Abs(($received - $target).TotalMinutes) -le 2
            })
            if ($byTime.Count -ge 1) { $candidates = $byTime }
        }
    }
    return @($candidates)
}

# The last resort: put the subject in Outlook's search box scoped to every
# folder and raise the window. Outlook's own search works on this profile even
# though AdvancedSearch does not, so a miss on our side is worth handing over.
function Show-SearchResults($ol, $ns, [string]$subject) {
    $query = 'subject:"' + ($subject -replace '"', '') + '"'
    $explorer = Read-Prop { $ol.ActiveExplorer() }
    if (-not $explorer) {
        $explorer = $ns.GetDefaultFolder(6).GetExplorer()
        $null = $explorer.Display()
        Start-Sleep -Milliseconds 1500
    }
    # OlSearchScope: 0 CurrentFolder, 1 CurrentFolderAndSubfolders,
    # 2 AllFolders, 3 AllOutlookItems. 2 is the one that lands on "All Outlook
    # Items"; 3 measured as a no-op that leaves the current folder up.
    $explorer.Search($query, 2)
    Start-Sleep -Milliseconds 500
    $null = Read-Prop { $explorer.Activate() }
    return $query
}

# --- Request ---------------------------------------------------------------

try {
    $req = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
    Emit @{ ok = $false; reason = 'bad_request'; message = $_.Exception.Message }
}

$action = if ($req.action) { [string]$req.action } else { 'resolve' }
if ($action -notin @('resolve', 'open', 'reply', 'reply_all', 'forward', 'save_attachments')) {
    Emit @{ ok = $false; reason = 'bad_request'; message = "unknown action '$action'" }
}

$hasRefs = -not [string]::IsNullOrWhiteSpace($req.entry_id)
$hasSubject = -not [string]::IsNullOrWhiteSpace($req.subject)
if (-not $hasRefs -and -not $hasSubject) {
    Emit @{ ok = $false; reason = 'bad_request'
            message = 'request carries neither entry_id nor subject' }
}

if ($action -eq 'save_attachments') {
    if (-not $hasRefs) {
        Emit @{ ok = $false; reason = 'bad_request'
                message = 'save_attachments needs an entry_id' }
    }
    if ([string]::IsNullOrWhiteSpace($req.dir) -or
        -not (Test-Path -LiteralPath ([string]$req.dir) -PathType Container)) {
        Emit @{ ok = $false; reason = 'bad_request'
                message = "save_attachments needs an existing 'dir'" }
    }
}

# --- Outlook ---------------------------------------------------------------

# Outlook is a single-instance COM server, so this attaches to a running classic
# instance or launches one. Which client is *default* is irrelevant: the
# Outlook.Application ProgID is registered by the classic Win32 client only, and
# New Outlook does not expose it -- so "fire up Outlook Classic" happens by
# construction rather than by detection.
try {
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
} catch {
    Emit @{ ok = $false; reason = 'no_outlook'; message = $_.Exception.Message }
}

$item = $null
$matchedBy = $null

if ($hasRefs) {
    # The one-argument GetItemFromID only searches the default store, so anything
    # in an archive or a shared mailbox resolves to nothing without the store_id.
    # That is why the macro writes both -- and why a failure here is reported
    # rather than retried by subject: with the store id supplied, it means the
    # message really has moved or been deleted.
    try {
        if ([string]::IsNullOrWhiteSpace($req.store_id)) {
            $item = $ns.GetItemFromID($req.entry_id)
        } else {
            $item = $ns.GetItemFromID($req.entry_id, $req.store_id)
        }
    } catch {
        Emit @{ ok = $false; reason = 'not_found'; message = $_.Exception.Message }
    }
    if (-not $item) {
        Emit @{ ok = $false; reason = 'not_found'; message = 'GetItemFromID returned nothing' }
    }
    $matchedBy = 'entry_id'
} else {
    $candidates = Select-Unique (Find-BySubject $ns ([string]$req.subject)) $req

    # Nothing, or more than one: both are "did not resolve uniquely", and
    # Outlook's own search is strictly better placed than we are to settle it.
    if (@($candidates).Count -ne 1) {
        try {
            $query = Show-SearchResults $ol $ns ([string]$req.subject)
        } catch {
            Emit @{ ok = $false; reason = 'com_error'; message = $_.Exception.Message }
        }
        Emit @{
            ok = $true; action = $action; matched_by = 'search_ui'
            query = $query; hits = @($candidates).Count
        }
    }

    $item = @($candidates)[0]
    $matchedBy = 'subject'
}

if (-not $item) {
    Emit @{ ok = $false; reason = 'not_found'; message = 'resolved to no item' }
}

$messageClass = Read-Prop { $item.MessageClass } ''
$subject      = Read-Prop { $item.Subject } ''
$when         = Read-Prop { $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') }
if (-not $when) { $when = Read-Prop { $item.Start.ToString('yyyy-MM-dd HH:mm') } '' }
$conversation = Read-Prop { $item.ConversationID } ''
$entryId      = Read-Prop { $item.EntryID } ([string]$req.entry_id)
$storeId      = Read-Prop { $item.Parent.StoreID } ([string]$req.store_id)
# PR_INTERNET_MESSAGE_ID -- absent on drafts and on some internal-only items.
$messageId    = Read-Prop {
    $item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001E')
} ''

# --- Act -------------------------------------------------------------------

# Meetings resolve and display fine but have no mail-shaped Reply/Forward, so
# say that rather than letting COM raise something opaque.
if ($action -in @('reply', 'reply_all', 'forward') -and $messageClass -notlike 'IPM.Note*') {
    Emit @{
        ok = $false; reason = 'unsupported_item'
        message = "'$action' is not available for an item of class '$messageClass'"
        subject = $subject; message_class = $messageClass
    }
}

if ($action -eq 'save_attachments') {
    # One bad attachment costs that attachment, not the rest: an embedded item
    # or a name the filesystem rejects is reported in 'failed' and the loop
    # goes on. $saved / $failed stay arrays through +=, so a single name still
    # serialises as a JSON list.
    $prefix = if ($req.prefix) { [string]$req.prefix } else { '' }
    $saved = @()
    $failed = @()
    try {
        foreach ($att in $item.Attachments) {
            $name = Read-Prop { [string]$att.FileName } ''
            if (-not $name) { continue }
            try {
                $att.SaveAsFile((Join-Path ([string]$req.dir) ($prefix + $name)))
                $saved += ($prefix + $name)
            } catch { $failed += $name }
        }
    } catch {
        Emit @{ ok = $false; reason = 'com_error'; message = $_.Exception.Message }
    }
    Emit @{
        ok = $true; action = $action; matched_by = $matchedBy
        subject = $subject; saved = @($saved); failed = @($failed)
    }
}

try {
    $compose = $null
    switch ($action) {
        'open'      { $null = $item.Display(); $null = Read-Prop { $item.GetInspector.Activate() } }
        'reply'     { $compose = $item.Reply() }
        'reply_all' { $compose = $item.ReplyAll() }
        'forward'   { $compose = $item.Forward() }
    }
    if ($compose) {
        $null = $compose.Display()
        # Without this the compose window routinely opens behind Sublime.
        $null = Read-Prop { $compose.GetInspector.Activate() }
    }
} catch {
    Emit @{ ok = $false; reason = 'com_error'; message = $_.Exception.Message }
}

Emit @{
    ok              = $true
    action          = $action
    matched_by      = $matchedBy
    entry_id        = $entryId
    store_id        = $storeId
    message_class   = $messageClass
    subject         = $subject
    received        = $when
    conversation_id = $conversation
    message_id      = $messageId
}
