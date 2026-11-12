Attribute VB_Name = "USER_CONFIG"
Option Explicit

Public Sub config(vaultPathToSaveFileTo As String, personNameStartChar As String, Optional emailFileNameStartChr As String, Optional emailTypeLink As String, Optional attachmentPath As String, Optional meetingFileNameStartChr As String, Optional meetingTypeLink As String, Optional trainingFileNameStartChr As String, Optional trainingTypeLink As String, Optional contactFileNameStartChr As String, Optional contactTypeLink As String, Optional sentFileNameStartChr As String, Optional personalVaultPathToSaveFileTo As String)
    
    '    Public Sub config(
    '                  vaultPathToSaveFileTo As String
    '                , personNameStartChar As String
    '                , Optional emailFileNameStartChr As String
    '                , Optional emailTypeLink As String
    '                , Optional meetingFileNameStartChr As String
    '                , Optional meetingTypeLink As String
    '                , Optional trainingFileNameStartChr As String
    '                , Optional trainingTypeLink As String
    '                , Optional contactFileNameStartChr As String
    '                , Optional contactTypeLink As String
    '   )
    
    '================================================'
    '====DECLARE=YOUR=FILE=PATH=TO=SAVE=FILES=TO====='
    '================================================'
    ' Make sure this is the absolute path where you want your files to be sent to
    ' !IMPORTANT! make sure you have a trailing backslash at the end of the path `\`
    vaultPathToSaveFileTo = "C:\Users\francis.dennig\OneDrive - United Nations Development Programme\UNDPnotes\emails\"
    ' Personal emails live in the vault's emails\ folder since 20261021, told apart by `personal: true`
    personalVaultPathToSaveFileTo = "C:\Users\francis.dennig\OneDrive - United Nations Development Programme\UNDPnotes\emails\"
    '================================================'
    ' This is what is added to the links inside the produced documents for people
    ' so it would look like:
    ' ex: "[[@Bryan Jenks]]"
    personNameStartChar = "@"
    '================================================'
    
    
    ' ^^^^^^^^^^^^^^^^^^^^^^^
    ' THE ABOVE ARE MANDATORY
    '
    ' THE BELOW ARE OPTIONAL
    ' VVVVVVVVVVVVVVVVVVVVVVV
    
    
    
    ' This is the beginning of the email file name and the H1 heading inside the file
    ' You need to leave a space in this if you want the name to appear like it does above
    ' ex: "E 2021-03-27 1347 Email Subject line here"
    emailFileNameStartChr = "Email_"
    '================================================'
    ' This is the beginning of the file name for emails you SEND with SendAndSaveEmail
    ' Keeping a distinct prefix lets you tell outgoing notes apart from received ones
    ' ex: "Sent_2021-03-27_1347 @Recipient - Subject"
    sentFileNameStartChr = "Sent_"
    '================================================'
    ' Path where attachments are saved when using ExtractEmailWithAttachments
    ' !IMPORTANT! make sure you have a trailing backslash at the end of the path `\`
    attachmentPath = "C:\Users\francis.dennig\OneDrive - United Nations Development Programme\UNDPnotes\emails\attachments\"
    '================================================'
    ' This is the beginning of the Meeting file name and the H1 heading inside the file
    ' You need to leave a space in this if you want the name to appear like it does above
    ' ex: "M 2021-03-27 1347 Meeting Subject Here"
    meetingFileNameStartChr = "Meeting_"
    '================================================'
    ' This is the beginning of the Training file name and the H1 heading inside the file
    ' You need to leave a space in this if you want the name to appear like it does above
    ' ex: "T 2021-03-27 1347 My Training Subject Here"
    trainingFileNameStartChr = "T "

    '================================================'
    ' This is the beginning of the contact file name and the H1 heading inside the file
    ' Do not leave a space if you want the result to look like:
    ' ex: "@Bryan Jenks"
    contactFileNameStartChr = "@"
    
    '================================================'
    ' These set the meta data links in the files so that there is a file named
    ' `+.md` that links to all emails. It's another way of sifting through your data
    ' You do not need to use these. To turn them off just leave a pair of empty quotes
    ' if left on they produce a field in the saved files that looks like:
    ' ex: "- `Type: ` [[+]]"
    emailTypeLink = "[[+]]"
    meetingTypeLink = "[[&]]"
    trainingTypeLink = "[[!]]"
    contactTypeLink = "[[@]]"

End Sub

' Settings for ExtractEmailsBatch only. Kept apart from config so that sub's
' already long optional signature does not grow again.
Public Sub batchConfig(stagingRoot As String, skipSenders As String, includeInboxSubfolders As Boolean, includeSentItems As Boolean, defaultDaysBack As Long)

    '================================================'
    ' Where a batch is staged before review: one folder per run underneath this.
    ' It must stay OUTSIDE the vault and outside OneDrive -- these are throwaway
    ' files, and inside the vault they would be synced and indexed before you had
    ' kept one. The Sublime side reads the same place from "staging_root" in
    ' EmailBatch.sublime-settings; change one and you must change the other.
    ' !IMPORTANT! make sure you have a trailing backslash at the end of the path `\`
    stagingRoot = Environ$("LOCALAPPDATA") & "\MarkdownNoteHarness\emailBatch\"
    '================================================'
    ' Mail from these senders is never exported. Fragments separated by ';', each
    ' matched anywhere in the sender's address, case-insensitively.
    skipSenders = "no-reply;noreply;no_reply;donotreply;do-not-reply;mailer-daemon;postmaster@"
    '================================================'
    ' The Inbox is always read. Its subfolders too?
    includeInboxSubfolders = True
    ' Sent Items too? Your own mail is as much part of a thread as the replies.
    includeSentItems = True
    '================================================'
    ' What the range prompt suggests the very first time, before there has been a
    ' batch to continue from.
    defaultDaysBack = 7

End Sub

