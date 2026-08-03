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

        {"action": "resolve|open|reply|reply_all|forward",
         "entry_id": "...", "store_id": "..."}

    Response: exactly one compact JSON object on stdout, always with exit 0.

        {"ok": true,  "action": "...", "matched_by": "entry_id",
         "entry_id": "...", "store_id": "...", "message_class": "IPM.Note",
         "subject": "...", "received": "...", "conversation_id": "...",
         "message_id": "..."}
        {"ok": false, "reason": "bad_request|no_outlook|not_found|
                                 unsupported_item|com_error", "message": "..."}

    Resolution is entry_id + store_id only -- one O(1) GetItemFromID. The
    message_id / time-window search tiers described in outlookReplyProposal.md
    are deliberately not built yet, so a moved or archived item (whose EntryID
    Outlook re-keys on the move) simply reports not_found.

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
function Read-Prop([scriptblock]$block) {
    try { & $block } catch { $null }
}

# --- Request ---------------------------------------------------------------

try {
    $req = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
    Emit @{ ok = $false; reason = 'bad_request'; message = $_.Exception.Message }
}

$action = if ($req.action) { [string]$req.action } else { 'resolve' }
if ($action -notin @('resolve', 'open', 'reply', 'reply_all', 'forward')) {
    Emit @{ ok = $false; reason = 'bad_request'; message = "unknown action '$action'" }
}

if ([string]::IsNullOrWhiteSpace($req.entry_id)) {
    Emit @{ ok = $false; reason = 'bad_request'; message = 'request carries no entry_id' }
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

# The one-argument GetItemFromID only searches the default store, so anything in
# an archive or a shared mailbox resolves to nothing without the store_id. That
# is why the macro writes both.
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

$messageClass = Read-Prop { $item.MessageClass }
$subject      = Read-Prop { $item.Subject }
$when         = Read-Prop { $item.ReceivedTime.ToString('yyyy-MM-dd HH:mm') }
if (-not $when) { $when = Read-Prop { $item.Start.ToString('yyyy-MM-dd HH:mm') } }
$conversation = Read-Prop { $item.ConversationID }
# PR_INTERNET_MESSAGE_ID -- absent on drafts and on some internal-only items.
$messageId    = Read-Prop {
    $item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001E')
}

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
    matched_by      = 'entry_id'
    entry_id        = $req.entry_id
    store_id        = $req.store_id
    message_class   = $messageClass
    subject         = $subject
    received        = $when
    conversation_id = $conversation
    message_id      = $messageId
}
