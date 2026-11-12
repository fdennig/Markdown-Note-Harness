Attribute VB_Name = "SaveEmail"

Option Explicit
'======================================================================================='

' Declare ShellExecute API for opening Obsidian links
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub ExtractEmail()
    ' Auto-dispatch based on selected item type
    Dim obj As Object
    Set obj = Application.ActiveExplorer.Selection.Item(1)
    Select Case obj.Class
        Case 43 ' olMail
            ExtractEmailCore False, False, False, ""
        Case 26 ' olAppointment
            ExtractMeetingCore
        Case 53 ' olMeetingItem (invitation in inbox)
            ExtractMeetingCore
        Case Else
            MsgBox "ExtractEmail: unsupported item type (Class " & obj.Class & ")."
    End Select
    Set obj = Nothing
End Sub

Sub ExtractEmailToReply()
    ExtractEmailCore True, False, False, ""
End Sub

Sub ExtractEmailWithAttachments()
    ExtractEmailCore False, True, False, ""
End Sub

Sub ExtractEmailPersonal()
    ExtractEmailCore False, False, True, ""
End Sub

Sub ExtractEmailHighPriority()
    ExtractEmailCore False, False, False, "high"
End Sub

Private Sub ExtractEmailCore(toReply As Boolean, saveAttachments As Boolean, Optional isPersonal As Boolean = False, Optional priority As String = "")

    Dim vaultPathToSaveFileTo As String
    Dim personalVaultPathToSaveFileTo As String
    Dim emailFileNameStartChr As String
    Dim emailTypeLink As String
    Dim personNameStartChar As String
    Dim attachmentPath As String
    config vaultPathToSaveFileTo, personNameStartChar, emailFileNameStartChr, emailTypeLink, attachmentPath, personalVaultPathToSaveFileTo:=personalVaultPathToSaveFileTo

    Dim targetVaultPath As String
    If isPersonal And Len(personalVaultPathToSaveFileTo) > 0 Then
        targetVaultPath = personalVaultPathToSaveFileTo
    Else
        targetVaultPath = vaultPathToSaveFileTo
    End If

    '================================================'
    ' Object holding variable
    Dim obj As Object
    ' Instantiate an Outlook Email Object
    Dim oMail As Outlook.MailItem
    ' If something breaks, skip to the end, tidy up and shut the door
    On Error GoTo EndClean:
    ' Establish the environment and selected items (emails)
    ' NOTE: selecting a conversation-view stack wont work
    '       you'll need to select one of the emails
    Dim fileName As String
    Dim resultString As String
    Dim bodyRow As Long
    Dim attPrefix As String
    Dim currentExplorer As Explorer
        Set currentExplorer = Application.ActiveExplorer
    Dim Selection As Selection
        Set Selection = currentExplorer.Selection
    ' For each email in the Selection
    ' Assigning email item to the `obj` holding variable
    For Each obj In Selection
        ' set the oMail object equal to that mail item
        Set oMail = obj
        ' Is it an Email?
        If oMail.Class <> 43 Then
          MsgBox "This code only works with Emails."
          GoTo EndClean: ' you broke it
        End If

        ' The note itself -- frontmatter, body and filename -- comes from
        ' BuildEmailNote, which ExtractEmailsBatch shares, so a note exported
        ' one at a time and a note exported in a batch cannot drift apart.
        resultString = BuildEmailNote(oMail, toReply, saveAttachments, isPersonal, priority, _
            personNameStartChar, emailFileNameStartChr, fileName, bodyRow, attPrefix)

    ' Save the result
    SaveAsUTF8 targetVaultPath & fileName, resultString

    ' Save attachments to disk if requested
    If saveAttachments And oMail.Attachments.Count > 0 Then
        Dim attPath As String
        attPath = attachmentPath
        Dim attItem As Attachment
        For Each attItem In oMail.Attachments
            attItem.SaveAsFile attPath & attPrefix & attItem.fileName
        Next attItem
    End If

    ' Open the saved markdown file in Sublime Text instead of Obsidian.
    ' Requires the 'subl' command line tool to be available in PATH.
    ' We pass the full file path as a quoted parameter (ShellExecute handles argument string separately).
    ' The ':row:col' position spec goes INSIDE the quotes: subl parses the
    ' trailing colon groups off the end, so the drive letter's colon is safe.
    ' CStr rather than bare '&' concatenation only to be explicit that no
    ' leading space (Str$'s) is wanted. Both are 1-based; the BOM SaveAsUTF8
    ' writes shifts neither.
    ShellExecute 0, "open", "subl", _
        """" & targetVaultPath & fileName & ":" & CStr(bodyRow) & ":1" & """", vbNullString, 1


    Next

    ' Normal completion. Jumping over the handler is what lets it report an
    ' error without false alarms: reaching the label below can then only mean
    ' something failed. Testing Err.Number there would not work -- the
    ' 'On Error Resume Next' regions above leave Err set on a successful run,
    ' and neither Resume Next nor a later 'On Error GoTo' clears it.
    Set obj = Nothing
    Set oMail = Nothing
    Exit Sub

EndClean:
    ' This used to be silent, which made every failure indistinguishable from
    ' the macro not being wired up at all -- no note, no message, nothing to
    ' search for. SendAndSaveEmail has always reported; this now matches it.
    MsgBox "ExtractEmail failed: " & Err.Description & _
           " (error " & Err.Number & ")" & vbCrLf & vbCrLf & _
           "No note was written. The file is saved only after the whole " & _
           "frontmatter is built, so there is no half-written note to clean up.", _
           vbExclamation, "ExtractEmail"
    Set obj = Nothing
    Set oMail = Nothing
End Sub

'---------------------------------------------------------------------------'
' One email as a note: returns the text, and hands back the filename, the row
' of the blank line under the frontmatter, and the attachment prefix.
'
' Shared by ExtractEmailCore (the selection, one note at a time, straight into
' the vault) and ExtractEmailsBatch (a date range, into a staging folder). It
' writes nothing and opens nothing; that is the caller's business.
'
' It has NO error handler of its own, on purpose: a failure here surfaces in
' the caller's handler, so ExtractEmailCore still reports through EndClean and
' the batch still counts the item as failed and moves on.
'---------------------------------------------------------------------------'
Private Function BuildEmailNote(oMail As Outlook.MailItem, toReply As Boolean, saveAttachments As Boolean, _
        isPersonal As Boolean, priority As String, personNameStartChar As String, _
        emailFileNameStartChr As String, ByRef fileName As String, ByRef bodyRow As Long, _
        ByRef attPrefix As String) As String

        Dim temporarySubjectLineString As String

        ' Keep the subject twice over: the raw one goes into the note, the
        ' cleaned one into the filename. These used to be a single string, so
        ' filename sanitising silently stripped ':' '?' '/' and friends out of
        ' the frontmatter as well, leaving a subject that could never be
        ' matched back against Outlook's own.
        Dim rawSubject As String
        rawSubject = oMail.Subject
        temporarySubjectLineString = rawSubject
        ' function call the name cleaner to remove any
        '    illegal characters from the subject line
        ReplaceCharsForFileName temporarySubjectLineString, ""
        ' Yank the received date-time to a holding variable

        ' Build Recipient string based on receipient collection
        Dim recips As Outlook.Recipients
            Set recips = oMail.Recipients
        Dim recip As Outlook.Recipient
        Dim result As String
        Dim recipString As String
            recipString = ""

        For Each recip In recips
            recipString = recipString & vbTab
            recipString = recipString & "- "
            recipString = recipString & formatName(recip.Name, personNameStartChar)
            recipString = recipString & vbCrLf
        Next
        ' Build the result file content to be sent to the mail item body
        ' Then save that mail item same as the meeting extractor
        Dim sender As String
            sender = formatName(oMail.sender, personNameStartChar)
        Dim dtDate As Date
            dtDate = oMail.ReceivedTime
        Dim resultString As String

        ' resultString = ""
        ' resultString = resultString & "# [[" & emailFileNameStartChr & Format(oMail.ReceivedTime, "yyyy-mm-dd hhnn") & " " & temporarySubjectLineString & "|" & temporarySubjectLineString & "]]"
        ' resultString = resultString & vbCrLf & vbCrLf & vbCrLf
        
        ' Start YAML frontmatter
        resultString = "---" & vbCrLf
        resultString = resultString & "toreply: " & IIf(toReply, "true", "") & vbCrLf
        resultString = resultString & "personal: " & IIf(isPersonal, "true", "") & vbCrLf
        resultString = resultString & "priority: " & priority & vbCrLf
        resultString = resultString & "class: email" & vbCrLf
        resultString = resultString & "subject: " & YamlQuoted(rawSubject) & vbCrLf
        resultString = resultString & "date: " & Format(oMail.ReceivedTime, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "from: """ & sender & """" & vbCrLf
        ' Add sender email address if available (placed once here)
        Dim senderEmail As String
        senderEmail = GetSenderSmtp(oMail)
        If Len(senderEmail) > 0 Then
            resultString = resultString & "from_email: """ & senderEmail & """" & vbCrLf
        End If
        ' Recipients: two keys, built in one pass over the collection.
        '   to:        EVERY recipient -- To, Cc and Bcc alike.
        '   to_email:  their SMTP addresses, one per line, in the SAME ORDER.
        '
        ' There is deliberately no cc: split. Cc is not used consistently enough
        ' to carry meaning -- the difference between addressing someone and
        ' copying them says more about the sender's habits than about who the
        ' message concerned -- so encoding it would add a key that invites false
        ' inferences. The Recipients collection is not even ordered To-then-Cc
        ' (measured: a live item with both Cc entries ahead of the To), so the
        ' order here is Outlook's, not a ranking.
        '
        ' Order and count are the whole contract on to_email:. The resolver
        ' pairs it with to: BY POSITION and only when the two lists are the same
        ' length, and its YAML reader discards empty items -- so one skipped or
        ' blank line does not cost one address, it throws away every address on
        ' the note and drops it back to matching on display names. That is why
        ' GetRecipSmtp always returns something.
        Dim toNames As String: toNames = ""
        Dim toAddrs As String: toAddrs = ""
        For Each recip In recips
            toNames = toNames & "  - " & YamlQuoted(formatName(recip.Name, personNameStartChar)) & vbCrLf
            toAddrs = toAddrs & "  - " & YamlQuoted(GetRecipSmtp(recip)) & vbCrLf
        Next
        resultString = resultString & YamlSeq("to", toNames)
        resultString = resultString & YamlSeq("to_email", toAddrs)
        resultString = resultString & "tags:" & vbCrLf
        ' Attachments list inside front matter
        Dim att As Attachment, attName As String
        attPrefix = IIf(saveAttachments, Format(dtDate, "yyyy-MM-dd_HHNN") & "_", "")
        If oMail.Attachments.Count > 0 Then
            resultString = resultString & "attachments:" & vbCrLf
            For Each att In oMail.Attachments
                attName = att.fileName
                attName = Replace(attName, """", "''")
                resultString = resultString & "  - """ & attPrefix & attName & """" & vbCrLf
            Next att
        Else
            resultString = resultString & "attachments: []" & vbCrLf
        End If

        ' --- Outlook back-references ----------------------------------------
        ' Handles on the live Outlook item, so the Sublime plugin can reply,
        ' reply-all or forward inside the original thread rather than build a
        ' mailto: that starts a new one.
        '
        ' These values are long and opaque and nobody reads them, so they go
        ' last and go in a sequence -- the same shape as 'attachments:', which
        ' folds to a single line and stays folded while the rest of the
        ' frontmatter is being read.
        '   entry_id + store_id  O(1) GetItemFromID. Both are needed: the
        '                        one-argument form searches only the default
        '                        store, so archives and shared mailboxes miss.
        '   message_id           stable forever, so it still resolves after the
        '                        item moves and its EntryID is re-keyed.
        '   conversation_id      not a lookup key on its own; it picks the right
        '                        thread out of several same-subject candidates.
        resultString = resultString & "outlook:" & vbCrLf
        resultString = resultString & "  - entry_id: " & GetEntryID(oMail) & vbCrLf
        resultString = resultString & "  - store_id: " & GetStoreID(oMail) & vbCrLf
        Dim internetMsgId As String
        internetMsgId = GetInternetMessageId(oMail)
        If Len(internetMsgId) > 0 Then
            resultString = resultString & "  - message_id: " & YamlQuoted(internetMsgId) & vbCrLf
        End If
        resultString = resultString & "  - conversation_id: " & GetConversationID(oMail) & vbCrLf

    ' End YAML block then blank line
    resultString = resultString & "---" & vbCrLf & vbCrLf
    ' Row of that blank line, for the 'subl file:row:col' spec below, so the
    ' caret lands after the frontmatter instead of at the top of the file.
    ' It has to be counted rather than known: the frontmatter's length varies
    ' with the recipient and attachment lists. The last Split element is the
    ' body line not yet written, so UBound is the blank line and UBound + 1
    ' would be the first line of the body.
    bodyRow = UBound(Split(resultString, vbCrLf))
    ' Add a Markdown task before the email body
    ' Add original email content
    resultString = resultString & ConvertLinksToMarkdown(oMail.Body)



        ' Build improved filename for fuzzy search:
        ' Format: <prefix><yyyy-MM-dd>_<HHNN> @Sender - Subject.md
        ' Ensures: single underscore between date/time, single @, normalized sender spacing.
        Dim dateToken As String: dateToken = Format(dtDate, "yyyy-MM-dd", vbUseSystemDayOfWeek, vbUseSystem)
        Dim timeToken As String: timeToken = Format(dtDate, "HHNN", vbUseSystemDayOfWeek, vbUseSystem)
        Dim senderToken As String: senderToken = sender
        On Error Resume Next
        ReplaceCharsForFileName senderToken, " "
        On Error GoTo 0
        senderToken = Replace(senderToken, "_", " ")
        senderToken = CompressWhitespace(Trim$(senderToken))
    ' Ensure exactly one leading @ before sender name in filename
    If Left$(senderToken, 1) = "@" Then senderToken = Mid$(senderToken, 2)
    senderToken = LTrim$(senderToken)
        ' Truncate subject for filename if excessively long (keep search-friendly, avoid huge entries)
        Dim subjectForFile As String
        subjectForFile = temporarySubjectLineString
        If Len(subjectForFile) > 60 Then
            subjectForFile = Left$(subjectForFile, 60)
        End If
    fileName = emailFileNameStartChr & dateToken & "_" & timeToken & " @" & senderToken & " - " & subjectForFile & ".md"

    BuildEmailNote = resultString
End Function

Private Sub ExtractMeetingCore()

    Dim vaultPathToSaveFileTo As String
    Dim meetingFileNameStartChr As String
    Dim personNameStartChar As String
    config vaultPathToSaveFileTo, personNameStartChar, meetingFileNameStartChr:=meetingFileNameStartChr

    Dim obj As Object
    Dim oAppt As Outlook.AppointmentItem
    On Error GoTo EndCleanMeeting:

    Dim currentExplorer As Explorer
        Set currentExplorer = Application.ActiveExplorer
    Dim Selection As Selection
        Set Selection = currentExplorer.Selection

    For Each obj In Selection
        ' Handle both calendar appointments (Class 26) and inbox meeting items (Class 53)
        If obj.Class = 53 Then
            Set oAppt = obj.GetAssociatedAppointment(False)
        ElseIf obj.Class = 26 Then
            Set oAppt = obj
        Else
            MsgBox "This code only works with Meetings/Appointments."
            GoTo EndCleanMeeting:
        End If

        ' Raw subject for the note, cleaned subject for the filename -- the
        ' same split as ExtractEmailCore
        Dim rawSubject As String
        Dim temporarySubjectLineString As String
        rawSubject = oAppt.Subject
        temporarySubjectLineString = rawSubject
        ReplaceCharsForFileName temporarySubjectLineString, ""

        ' Build required attendees list
        Dim recips As Outlook.Recipients
            Set recips = oAppt.Recipients
        Dim recip As Outlook.Recipient
        Dim reqAttendees As String: reqAttendees = ""
        Dim optAttendees As String: optAttendees = ""

        For Each recip In recips
            If recip.Type = olRequired Then
                reqAttendees = reqAttendees & "  - """ & formatName(recip.Name, personNameStartChar) & """" & vbCrLf
            ElseIf recip.Type = olOptional Then
                optAttendees = optAttendees & "  - """ & formatName(recip.Name, personNameStartChar) & """" & vbCrLf
            End If
        Next

        ' Build YAML frontmatter
        Dim resultString As String
        resultString = "---" & vbCrLf
        resultString = resultString & "class: meeting" & vbCrLf
        resultString = resultString & "subject: " & YamlQuoted(rawSubject) & vbCrLf
        resultString = resultString & "date: " & Format(oAppt.Start, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "end: " & Format(oAppt.End, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "duration: " & oAppt.Duration & vbCrLf

        Dim locationStr As String
        locationStr = oAppt.Location
        If Len(locationStr) > 0 Then
            locationStr = Replace(locationStr, """", "''")
            resultString = resultString & "location: """ & locationStr & """" & vbCrLf
        Else
            resultString = resultString & "location: " & vbCrLf
        End If

        resultString = resultString & "organizer: """ & formatName(oAppt.Organizer, personNameStartChar) & """" & vbCrLf

        ' Attendees
        If Len(reqAttendees) > 0 Then
            resultString = resultString & "attendees:" & vbCrLf & reqAttendees
        Else
            resultString = resultString & "attendees: []" & vbCrLf
        End If

        If Len(optAttendees) > 0 Then
            resultString = resultString & "optional_attendees:" & vbCrLf & optAttendees
        Else
            resultString = resultString & "optional_attendees: []" & vbCrLf
        End If

        resultString = resultString & "tags:" & vbCrLf

        ' Attachments
        Dim att As Attachment, attName As String
        If oAppt.Attachments.Count > 0 Then
            resultString = resultString & "attachments:" & vbCrLf
            For Each att In oAppt.Attachments
                attName = att.fileName
                attName = Replace(attName, """", "''")
                resultString = resultString & "  - """ & attName & """" & vbCrLf
            Next att
        Else
            resultString = resultString & "attachments: []" & vbCrLf
        End If

        ' Outlook back-reference; same foldable sequence as ExtractEmailCore.
        ' A meeting cannot be replied to, but this is what lets the plugin open
        ' the original appointment in Outlook.
        resultString = resultString & "outlook:" & vbCrLf
        resultString = resultString & "  - entry_id: " & GetEntryID(oAppt) & vbCrLf
        resultString = resultString & "  - store_id: " & GetStoreID(oAppt) & vbCrLf

        resultString = resultString & "---" & vbCrLf & vbCrLf
        ' Caret row for subl; see the note in ExtractEmailCore.
        Dim bodyRow As Long
        bodyRow = UBound(Split(resultString, vbCrLf))

        ' Body
        resultString = resultString & ConvertLinksToMarkdown(oAppt.Body)

        ' Build filename: M <yyyy-MM-dd>_<HHNN> - Subject.md
        Dim dtDate As Date: dtDate = oAppt.Start
        Dim dateToken As String: dateToken = Format(dtDate, "yyyy-MM-dd")
        Dim timeToken As String: timeToken = Format(dtDate, "HHNN")
        Dim subjectForFile As String: subjectForFile = temporarySubjectLineString
        If Len(subjectForFile) > 60 Then subjectForFile = Left$(subjectForFile, 60)

        Dim fileName As String
        fileName = meetingFileNameStartChr & dateToken & "_" & timeToken & " - " & subjectForFile & ".md"

        ' Save
        SaveAsUTF8 vaultPathToSaveFileTo & fileName, resultString

        ' Open in Sublime Text, caret after the frontmatter
        ShellExecute 0, "open", "subl", _
            """" & vaultPathToSaveFileTo & fileName & ":" & CStr(bodyRow) & ":1" & """", vbNullString, 1

    Next

    ' Same shape as ExtractEmailCore: fall out before the handler so that
    ' reaching it means failure and nothing else.
    Set obj = Nothing
    Set oAppt = Nothing
    Exit Sub

EndCleanMeeting:
    MsgBox "ExtractMeeting failed: " & Err.Description & _
           " (error " & Err.Number & ")" & vbCrLf & vbCrLf & _
           "No note was written.", vbExclamation, "ExtractMeeting"
    Set obj = Nothing
    Set oAppt = Nothing
End Sub

' Collapse multiple spaces into single spaces
Private Function CompressWhitespace(ByVal s As String) As String
    Dim i As Long, ch As String * 1, prevSpace As Boolean, outStr As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Then
            If Not prevSpace Then outStr = outStr & ch
            prevSpace = True
        Else
            outStr = outStr & ch
            prevSpace = False
        End If
    Next i
    CompressWhitespace = outStr
End Function

Private Function GetSenderSmtp(oMail As Outlook.MailItem) As String
    On Error GoTo Fail
    Dim addr As String
    If oMail.SenderEmailType = "SMTP" Then
        addr = oMail.SenderEmailAddress
    ElseIf oMail.SenderEmailType = "EX" Then
        Dim exchUser As Outlook.ExchangeUser
        On Error Resume Next
        Set exchUser = oMail.sender.GetExchangeUser
        On Error GoTo Fail
        If Not exchUser Is Nothing Then
            addr = exchUser.PrimarySmtpAddress
        End If
        If Len(addr) = 0 Then
            Dim pa As Outlook.PropertyAccessor
            Set pa = oMail.PropertyAccessor
            ' PR_SENDER_SMTP_ADDRESS -- the SENDER's address. This used to read
            ' PR_SMTP_ADDRESS (0x39FE001E), which belongs to the recipient
            ' table: on a mail item it comes back useless, so this fallback
            ' silently produced nothing every time GetExchangeUser failed.
            addr = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
        End If
    End If
    GetSenderSmtp = LCase$(addr)
    Exit Function
Fail:
    GetSenderSmtp = ""
End Function

Private Function GetRecipSmtp(recip As Outlook.Recipient) As String
    ' The SMTP address of one recipient, and NEVER an empty string -- see the
    ' recipient block in ExtractEmailCore for why a blank is worse than a wrong
    ' one.
    '
    ' Unlike the sender case above, PR_SMTP_ADDRESS is the correct property
    ' here: it is the recipient table's own, which is exactly what an
    ' AddressEntry reached through a Recipient is.
    Dim addr As String
    Dim ae As Outlook.AddressEntry

    On Error Resume Next
    Set ae = recip.AddressEntry
    If Not ae Is Nothing Then
        If ae.Type = "EX" Then
            ' Never chain off GetExchangeUser. "EX" says only that the entry
            ' came from Exchange, not that it is a person: a distribution list,
            ' a room or equipment mailbox, a public folder, and any entry the
            ' address book could not resolve all return Nothing -- and
            ' .PrimarySmtpAddress on Nothing is run-time error 91. The
            ' 'On Error Resume Next' above does not save it either, because
            ' 'Break on All Errors' in the VBE breaks inside a Resume Next
            ' region, so a single group in the To line stops the macro dead.
            Dim exchUser As Outlook.ExchangeUser
            Dim exchList As Outlook.ExchangeDistributionList
            Set exchUser = ae.GetExchangeUser
            If Not exchUser Is Nothing Then
                addr = exchUser.PrimarySmtpAddress
            Else
                ' A group has an SMTP address of its own, and it is the honest
                ' answer for that position in the list -- better than the X500
                ' DN the fallbacks below would otherwise leave there.
                Set exchList = ae.GetExchangeDistributionList
                If Not exchList Is Nothing Then
                    addr = exchList.PrimarySmtpAddress
                End If
            End If
            If Len(addr) = 0 Then
                addr = ae.PropertyAccessor.GetProperty( _
                    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
            End If
        Else
            addr = ae.Address
        End If
    End If
    ' Whatever the item itself carries, if the address book could not resolve
    ' the entry. For an Exchange recipient this can be an X500 DN, which
    ' matches no person note -- harmless, because an address that matches
    ' nothing does not stop the resolver falling back to the display name.
    If Len(addr) = 0 Then addr = recip.Address
    ' Last resort: hold the position in the list. Non-empty so the YAML reader
    ' keeps it and the two lengths still agree; not address-shaped, so it cannot
    ' collide with a real person note.
    If Len(addr) = 0 Then addr = "unknown"
    On Error GoTo 0

    GetRecipSmtp = LCase$(Trim$(addr))
End Function

Private Function YamlSeq(key As String, items As String) As String
    ' 'key:' followed by the items, or 'key: []' when there are none -- the
    ' convention attachments: and optional_attendees: already use, and the one
    ' the resolver reads as an empty list rather than as a missing key.
    If Len(items) > 0 Then
        YamlSeq = key & ":" & vbCrLf & items
    Else
        YamlSeq = key & ": []" & vbCrLf
    End If
End Function

'---------------------------------------------------------------------------'
' Outlook back-references for the Sublime plugin.
' Each one is defensive on purpose: a missing back-reference costs the plugin
' one search, whereas an unhandled error here would cost the whole note.
'---------------------------------------------------------------------------'
Private Function GetEntryID(oItem As Object) As String
    On Error GoTo Fail
    GetEntryID = oItem.EntryID
    Exit Function
Fail:
    GetEntryID = ""
End Function

Private Function GetStoreID(oItem As Object) As String
    On Error GoTo Fail
    ' oItem.Parent is the containing folder, and its StoreID is the second
    ' argument GetItemFromID needs to reach outside the default store.
    GetStoreID = oItem.Parent.StoreID
    Exit Function
Fail:
    GetStoreID = ""
End Function

Private Function GetConversationID(oItem As Object) As String
    On Error GoTo Fail
    GetConversationID = oItem.ConversationID
    Exit Function
Fail:
    GetConversationID = ""
End Function

Private Function GetInternetMessageId(oMail As Outlook.MailItem) As String
    On Error GoTo Fail
    ' PR_INTERNET_MESSAGE_ID -- the RFC 5322 Message-ID header. Absent on
    ' drafts and on some internal-only items, hence the caller's Len() check.
    GetInternetMessageId = oMail.PropertyAccessor.GetProperty( _
        "http://schemas.microsoft.com/mapi/proptag/0x1035001E")
    Exit Function
Fail:
    GetInternetMessageId = ""
End Function

Private Function ConvertLinksToMarkdown(bodyText As String) As String
    Dim result As String
    result = bodyText
    
    ' Convert <http...> patterns to ["link"](http...)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "<(http[^>]*)>"
    result = regex.Replace(result, "[link](<$1>)")
    
    ConvertLinksToMarkdown = result
End Function

'======================================================================================='
' SEND + SAVE
'  Run this from an OPEN COMPOSE WINDOW (Inspector). It writes the outgoing message as a
'  Markdown note (same layout as ExtractEmail) and then sends it, so one click both files
'  and sends the email. Hook it to a button on the compose window's Quick Access Toolbar.
'======================================================================================='
Sub SendAndSaveEmail()

    Dim vaultPathToSaveFileTo As String
    Dim personNameStartChar As String
    Dim sentFileNameStartChr As String
    config vaultPathToSaveFileTo, personNameStartChar, sentFileNameStartChr:=sentFileNameStartChr
    If Len(sentFileNameStartChr) = 0 Then sentFileNameStartChr = "Sent_"

    On Error GoTo EndCleanSend

    ' Grab the message currently being composed
    If TypeName(Application.ActiveInspector) = "Nothing" Then
        MsgBox "Open the email you want to send in its own window first, then run SendAndSaveEmail."
        Exit Sub
    End If

    Dim oMail As Outlook.MailItem
    Set oMail = Application.ActiveInspector.CurrentItem
    If oMail.Class <> 43 Then
        MsgBox "This only works with email messages."
        Exit Sub
    End If

    ' Resolve names/addresses before we read them
    oMail.Recipients.ResolveAll

    ' Saving assigns a ConversationID to a fresh draft, and that is the only
    ' back-reference an outgoing note can carry: .Send invalidates this item's
    ' EntryID (the message moves to Sent Items and is re-keyed there) and the
    ' Message-ID is stamped by the transport, so neither exists yet. Sent notes
    ' therefore resolve through the plugin's search fallback, not GetItemFromID.
    On Error Resume Next
    oMail.Save
    On Error GoTo EndCleanSend
    Dim conversationId As String
    conversationId = GetConversationID(oMail)

    ' --- capture every field BEFORE sending (the object is invalid after .Send) ---
    Dim rawSubject As String
    Dim temporarySubjectLineString As String
    rawSubject = oMail.Subject
    temporarySubjectLineString = rawSubject
    ReplaceCharsForFileName temporarySubjectLineString, ""

    ' A draft has no SentOn yet, so stamp the moment we send
    Dim dtDate As Date
    dtDate = Now

    Dim sender As String
    sender = formatName(Application.Session.CurrentUser.Name, personNameStartChar)

    Dim recips As Outlook.Recipients
        Set recips = oMail.Recipients
    Dim recip As Outlook.Recipient

    ' --- build YAML frontmatter (mirrors ExtractEmailCore) ---
    Dim resultString As String
    resultString = "---" & vbCrLf
    resultString = resultString & "toreply: " & vbCrLf
    resultString = resultString & "personal: " & vbCrLf
    resultString = resultString & "priority: " & vbCrLf
    resultString = resultString & "class: email" & vbCrLf
    resultString = resultString & "sent: true" & vbCrLf
    resultString = resultString & "subject: " & YamlQuoted(rawSubject) & vbCrLf
    resultString = resultString & "date: " & Format(dtDate, "yyyy-MM-dd HH:mm") & vbCrLf
    resultString = resultString & "from: """ & sender & """" & vbCrLf

    Dim myEmail As String
    myEmail = GetCurrentUserSmtp()
    If Len(myEmail) > 0 Then
        resultString = resultString & "from_email: """ & myEmail & """" & vbCrLf
    End If

    ' Recipients -- the same two keys and the same positional contract as
    ' ExtractEmailCore; see the comment there. Bcc recipients are real on an
    ' outgoing note (they are not on received mail) and land in to: with
    ' everyone else, which is also the only place they are recorded.
    Dim toNames As String: toNames = ""
    Dim toAddrs As String: toAddrs = ""
    For Each recip In recips
        toNames = toNames & "  - " & YamlQuoted(formatName(recip.Name, personNameStartChar)) & vbCrLf
        toAddrs = toAddrs & "  - " & YamlQuoted(GetRecipSmtp(recip)) & vbCrLf
    Next
    resultString = resultString & YamlSeq("to", toNames)
    resultString = resultString & YamlSeq("to_email", toAddrs)

    resultString = resultString & "tags:" & vbCrLf

    ' Attachments
    Dim att As Attachment, attName As String
    If oMail.Attachments.Count > 0 Then
        resultString = resultString & "attachments:" & vbCrLf
        For Each att In oMail.Attachments
            attName = Replace(att.fileName, """", "''")
            resultString = resultString & "  - """ & attName & """" & vbCrLf
        Next att
    Else
        resultString = resultString & "attachments: []" & vbCrLf
    End If

    ' Outlook back-reference; see the block in ExtractEmailCore. A sent note
    ' can only ever carry this one, but it keeps the same sequence shape so
    ' there is a single thing to fold and a single thing to read.
    resultString = resultString & "outlook:" & vbCrLf
    resultString = resultString & "  - conversation_id: " & conversationId & vbCrLf

    resultString = resultString & "---" & vbCrLf & vbCrLf
    ' Caret row for subl; see the note in ExtractEmailCore.
    Dim bodyRow As Long
    bodyRow = UBound(Split(resultString, vbCrLf))
    resultString = resultString & ConvertLinksToMarkdown(oMail.Body)

    ' --- filename: <prefix><yyyy-MM-dd>_<HHNN> @<first recipient> - <subject>.md ---
    ' Outgoing mail is keyed on the recipient (the sender is always you)
    Dim dateToken As String: dateToken = Format(dtDate, "yyyy-MM-dd")
    Dim timeToken As String: timeToken = Format(dtDate, "HHNN")

    Dim toToken As String
    If recips.Count > 0 Then
        toToken = recips.Item(1).Name
    Else
        toToken = "no recipient"
    End If
    On Error Resume Next
    ReplaceCharsForFileName toToken, " "
    On Error GoTo EndCleanSend
    toToken = CompressWhitespace(Trim$(Replace(toToken, "_", " ")))
    If Left$(toToken, 1) = "@" Then toToken = LTrim$(Mid$(toToken, 2))

    Dim subjectForFile As String
    subjectForFile = temporarySubjectLineString
    If Len(subjectForFile) > 60 Then subjectForFile = Left$(subjectForFile, 60)

    Dim fileName As String
    fileName = sentFileNameStartChr & dateToken & "_" & timeToken & " @" & toToken & " - " & subjectForFile & ".md"

    ' --- save the note FIRST, then send the mail ---
    SaveAsUTF8 vaultPathToSaveFileTo & fileName, resultString
    oMail.Send

    ' Open the saved markdown file in Sublime Text (same as ExtractEmail).
    ' Requires the 'subl' command line tool to be available in PATH.
    ShellExecute 0, "open", "subl", _
        """" & vaultPathToSaveFileTo & fileName & ":" & CStr(bodyRow) & ":1" & """", vbNullString, 1

    Exit Sub
EndCleanSend:
    MsgBox "SendAndSaveEmail failed: " & Err.Description & vbCrLf & _
           "(The email was NOT sent if you see this before it left.)"
    Set oMail = Nothing
End Sub

' SMTP address of the signed-in user, for the from_email field on sent notes
Private Function GetCurrentUserSmtp() As String
    On Error GoTo Fail
    Dim ae As Outlook.AddressEntry
    Set ae = Application.Session.CurrentUser.AddressEntry
    If ae.Type = "EX" Then
        ' Same trap as GetRecipSmtp: GetExchangeUser can return Nothing, and
        ' .PrimarySmtpAddress on Nothing is error 91 -- here it would be caught
        ' by the handler below and silently cost the sent note its from_email.
        Dim exchUser As Outlook.ExchangeUser
        Set exchUser = ae.GetExchangeUser
        If Not exchUser Is Nothing Then
            GetCurrentUserSmtp = LCase$(exchUser.PrimarySmtpAddress)
        Else
            GetCurrentUserSmtp = LCase$(ae.Address)
        End If
    Else
        GetCurrentUserSmtp = LCase$(ae.Address)
    End If
    Exit Function
Fail:
    GetCurrentUserSmtp = ""
End Function

'======================================================================================='
' BATCH EXPORT
'  ExtractEmailsBatch writes EVERY email received or sent in a date range to a staging
'  folder OUTSIDE the vault and hands that folder to Sublime, where EmailBatch.py drops
'  what is already filed, proposes a wiki link for the rest, and lets you accept, replace
'  or discard each one. Only an email you accept there is ever written into the vault.
'
'  The notes are built by BuildEmailNote -- the same code ExtractEmail runs -- so a note
'  that arrives through a batch is indistinguishable from one exported by hand.
'
'  The range prompt takes:
'      7                          the last 7 days
'      2026-11-01                 from that day until now
'      2026-11-01 14:30           from that minute until now
'      2026-11-01..2026-11-07     between two days, both included
'  and offers the end of the previous batch as its default, so pressing Enter means
'  "everything since last time". An overlap costs nothing: the plugin drops what is
'  already in the vault and what you discarded before.
'
'  Settings (staging folder, senders to skip, which folders) are in batchConfig in
'  USER_CONFIG.
'======================================================================================='
Sub ExtractEmailsBatch()

    Dim vaultPathToSaveFileTo As String
    Dim personNameStartChar As String
    Dim emailFileNameStartChr As String
    config vaultPathToSaveFileTo, personNameStartChar, emailFileNameStartChr

    Dim stagingRoot As String
    Dim skipSenders As String
    Dim includeInboxSubfolders As Boolean
    Dim includeSentItems As Boolean
    Dim defaultDaysBack As Long
    batchConfig stagingRoot, skipSenders, includeInboxSubfolders, includeSentItems, defaultDaysBack

    On Error GoTo EndCleanBatch

    ' --- the range -------------------------------------------------------------
    Dim suggestion As String
    suggestion = ReadLastBatchEnd(stagingRoot)
    If Len(suggestion) = 0 Then suggestion = Format(Date - defaultDaysBack, "yyyy-mm-dd")

    Dim answer As String
    answer = InputBox("Export every email received or sent since:" & vbCrLf & vbCrLf & _
        "    7" & vbTab & vbTab & vbTab & "the last 7 days" & vbCrLf & _
        "    2026-11-01" & vbTab & vbTab & "from that day until now" & vbCrLf & _
        "    2026-11-01 14:30" & vbTab & "from that minute until now" & vbCrLf & _
        "    2026-11-01..2026-11-07" & vbTab & "both days included", _
        "ExtractEmailsBatch", suggestion)
    If Len(Trim$(answer)) = 0 Then Exit Sub        ' Cancel, or an emptied box

    Dim fromDate As Date
    Dim toDate As Date
    If Not ParseBatchRange(answer, fromDate, toDate) Then
        MsgBox "Could not read '" & answer & "' as a range." & vbCrLf & vbCrLf & _
               "Use a number of days, yyyy-mm-dd, yyyy-mm-dd hh:mm, or " & _
               "yyyy-mm-dd..yyyy-mm-dd.", vbExclamation, "ExtractEmailsBatch"
        Exit Sub
    End If

    ' --- the staging folder ----------------------------------------------------
    ' One folder per run, named for the moment it started, under a root that is
    ' NOT in OneDrive and NOT in the vault: a week of mail is a few hundred
    ' throwaway files, and inside the vault they would be synced, indexed by
    ' Sublime and offered as link targets before you had kept a single one.
    Dim batchPath As String
    batchPath = stagingRoot & Format(Now, "yyyymmdd_hhnnss") & "\"
    EnsureFolder batchPath

    ' --- the export ------------------------------------------------------------
    Dim nSaved As Long
    Dim nSkipped As Long
    Dim nFailed As Long
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    BatchFolder ns.GetDefaultFolder(olFolderInbox), includeInboxSubfolders, "ReceivedTime", _
        fromDate, toDate, skipSenders, batchPath, personNameStartChar, emailFileNameStartChr, _
        nSaved, nSkipped, nFailed
    If includeSentItems Then
        BatchFolder ns.GetDefaultFolder(olFolderSentMail), False, "SentOn", _
            fromDate, toDate, skipSenders, batchPath, personNameStartChar, emailFileNameStartChr, _
            nSaved, nSkipped, nFailed
    End If

    If nSaved = 0 Then
        RemoveFolderIfEmpty batchPath
        MsgBox "No emails to export between " & Format(fromDate, "yyyy-mm-dd hh:nn") & _
               " and " & Format(toDate, "yyyy-mm-dd hh:nn") & "." & vbCrLf & vbCrLf & _
               nSkipped & " skipped by the sender filter, " & nFailed & " could not be read.", _
               vbInformation, "ExtractEmailsBatch"
        Set ns = Nothing
        Exit Sub
    End If

    ' Where the next batch starts by default. Never later than now, so a range
    ' that ends in the future does not skip mail that has yet to arrive.
    WriteLastBatchEnd stagingRoot, IIf(toDate > Now, Now, toDate)

    ' --- hand over to Sublime ---------------------------------------------------
    ' subl --command "email_batch_prepare {\"dir\": \"C:/.../20261112_101500/\"}"
    ' Forward slashes on purpose: the path travels inside a JSON string inside a
    ' quoted argument, and a backslash would need escaping at both levels. If
    ' Sublime is not running this starts it, and the command may then arrive
    ' before the plugins have loaded -- "Email batch: Review" in the command
    ' palette picks the same folder up, so nothing is lost either way.
    ShellExecute 0, "open", "subl", _
        "--command ""email_batch_prepare {\""dir\"": \""" & Replace(batchPath, "\", "/") & "\""}""", _
        vbNullString, 1

    ' Silent on success: Sublime's status bar reports the counts. A failure is
    ' worth a click, because a mail that could not be exported is a mail that
    ' will never be offered for review.
    If nFailed > 0 Then
        MsgBox nSaved & " emails exported; " & nFailed & " could not be read and were left out " & _
               "(encrypted mail, or a subject that makes an invalid filename)." & vbCrLf & vbCrLf & _
               "Export those one at a time with ExtractEmail to see the error.", _
               vbExclamation, "ExtractEmailsBatch"
    End If

    ' Normal completion; see ExtractEmailCore for why the handler is jumped over.
    Set ns = Nothing
    Exit Sub

EndCleanBatch:
    MsgBox "ExtractEmailsBatch failed: " & Err.Description & " (error " & Err.Number & ")" & _
           vbCrLf & vbCrLf & "Whatever was exported before the failure is in" & vbCrLf & _
           batchPath & vbCrLf & "and 'Email batch: Review' in Sublime will pick it up.", _
           vbExclamation, "ExtractEmailsBatch"
    Set ns = Nothing
End Sub

' Every mail item in one folder that falls inside the range, and optionally in
' its subfolders. The Restrict filter is deliberately coarse (see below); the
' exact test is made per item in BatchOne.
Private Sub BatchFolder(fld As Outlook.Folder, recurse As Boolean, dateField As String, _
        fromDate As Date, toDate As Date, skipSenders As String, batchPath As String, _
        personNameStartChar As String, emailFileNameStartChr As String, _
        ByRef nSaved As Long, ByRef nSkipped As Long, ByRef nFailed As Long)

    ' A Jet date literal has to be in the machine's own short-date format and may
    ' not carry seconds -- 'ddddd h:nn AMPM' is the documented way to write one.
    ' Because seconds are dropped, the upper bound is pushed out by a minute and
    ' BatchOne makes the precise comparison.
    Dim flt As String
    flt = "[" & dateField & "] >= '" & Format(fromDate, "ddddd h:nn AMPM") & "'" & _
          " AND [" & dateField & "] <= '" & Format(toDate + TimeSerial(0, 1, 0), "ddddd h:nn AMPM") & "'"

    Dim found As Outlook.Items
    Set found = fld.Items.Restrict(flt)

    Dim itm As Object
    Dim oMail As Outlook.MailItem
    For Each itm In found
        ' 43 = olMail. Meeting requests (53), read receipts and delivery reports
        ' (46) share the Inbox and are none of them correspondence.
        If itm.Class = 43 Then
            Set oMail = itm
            Select Case BatchOne(oMail, fromDate, toDate, skipSenders, batchPath, _
                                 personNameStartChar, emailFileNameStartChr)
                Case 1: nSaved = nSaved + 1
                Case 0: nSkipped = nSkipped + 1
                Case Else: nFailed = nFailed + 1
            End Select
        End If
        DoEvents        ' keep Outlook's window alive through a long range
    Next

    If recurse Then
        Dim subFld As Outlook.Folder
        For Each subFld In fld.Folders
            BatchFolder subFld, True, dateField, fromDate, toDate, skipSenders, batchPath, _
                personNameStartChar, emailFileNameStartChr, nSaved, nSkipped, nFailed
        Next
    End If
End Sub

' One mail item into the staging folder. 1 = written, 0 = deliberately skipped,
' -1 = failed. It has its own handler so that one unreadable item -- encrypted
' mail is the usual one -- costs that item and not the batch.
Private Function BatchOne(oMail As Outlook.MailItem, fromDate As Date, toDate As Date, _
        skipSenders As String, batchPath As String, personNameStartChar As String, _
        emailFileNameStartChr As String) As Long

    On Error GoTo Fail

    Dim stamp As Date
    stamp = oMail.ReceivedTime
    If stamp < fromDate Or stamp >= toDate Then
        BatchOne = 0
        Exit Function
    End If
    ' Out-of-office replies and other rule-generated mail: IPM.Note.Rules.*
    If LCase$(oMail.MessageClass) Like "ipm.note.rules.*" Then
        BatchOne = 0
        Exit Function
    End If
    If IsSkippedSender(GetSenderSmtp(oMail), skipSenders) Then
        BatchOne = 0
        Exit Function
    End If

    ' No flags and no attachments: to-reply, priority and personal are set in
    ' Sublime during review, and an attachment worth keeping is still one
    ' ExtractEmailWithAttachments away, by F12 from the note.
    Dim fileName As String
    Dim bodyRow As Long
    Dim attPrefix As String
    Dim noteText As String
    noteText = BuildEmailNote(oMail, False, False, False, "", personNameStartChar, _
        emailFileNameStartChr, fileName, bodyRow, attPrefix)

    SaveAsUTF8 batchPath & UniqueStagedName(batchPath, fileName), noteText
    BatchOne = 1
    Exit Function
Fail:
    BatchOne = -1
End Function

' True when the sender's address contains any of the ';'-separated fragments.
Private Function IsSkippedSender(ByVal addr As String, ByVal skipSenders As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim fragment As String
    If Len(addr) = 0 Or Len(skipSenders) = 0 Then Exit Function
    addr = LCase$(addr)
    parts = Split(LCase$(skipSenders), ";")
    For i = LBound(parts) To UBound(parts)
        fragment = Trim$(parts(i))
        If Len(fragment) > 0 Then
            If InStr(addr, fragment) > 0 Then
                IsSkippedSender = True
                Exit Function
            End If
        End If
    Next
End Function

' Two mails from one sender in one minute with one subject would otherwise
' overwrite each other: SaveAsUTF8 saves with adSaveCreateOverWrite.
' FileSystemObject rather than Dir$, which cannot see a name holding characters
' outside the ANSI code page -- an en dash in a subject is enough.
Private Function UniqueStagedName(batchPath As String, fileName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim stem As String
    Dim n As Long
    stem = Left$(fileName, Len(fileName) - 3)       ' without ".md"
    UniqueStagedName = fileName
    n = 1
    Do While fso.FileExists(batchPath & UniqueStagedName)
        n = n + 1
        UniqueStagedName = stem & " (" & n & ").md"
    Loop
    Set fso = Nothing
End Function

' "7", "2026-11-01", "2026-11-01 14:30" or "2026-11-01..2026-11-07" into a
' half-open range [fromDate, toDate). False when it is none of those.
Private Function ParseBatchRange(ByVal s As String, ByRef fromDate As Date, ByRef toDate As Date) As Boolean
    Dim parts() As String
    s = Trim$(s)
    toDate = Now
    If Len(s) >= 1 And Len(s) <= 3 And Not (s Like "*[!0-9]*") Then
        fromDate = Date - CLng(s)
        ParseBatchRange = True
        Exit Function
    End If
    If InStr(s, "..") > 0 Then
        parts = Split(s, "..")
        If UBound(parts) <> 1 Then Exit Function
        If Not ParseStamp(Trim$(parts(0)), fromDate, False) Then Exit Function
        If Not ParseStamp(Trim$(parts(1)), toDate, True) Then Exit Function
    Else
        If Not ParseStamp(s, fromDate, False) Then Exit Function
    End If
    ParseBatchRange = (fromDate < toDate)
End Function

' yyyy-mm-dd, optionally followed by " hh:mm". Read by position and never by
' CDate, whose reading of 2026-11-01 depends on the machine's regional settings.
' A bare END day is included whole, i.e. it runs to the following midnight.
Private Function ParseStamp(ByVal s As String, ByRef d As Date, endOfDay As Boolean) As Boolean
    On Error GoTo Fail
    If Not (s Like "####-##-##*") Then Exit Function
    d = DateSerial(CInt(Mid$(s, 1, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    If Len(s) > 10 Then
        If Not (Mid$(s, 11) Like " ##:##") Then Exit Function
        d = d + TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), 0)
    ElseIf endOfDay Then
        d = d + 1
    End If
    ParseStamp = True
    Exit Function
Fail:
    ParseStamp = False
End Function

' The end of the previous batch as "yyyy-mm-dd hh:mm", or "" when there has not
' been one. It is the range prompt's default, nothing more: deleting the file
' only means the next prompt suggests defaultDaysBack again.
Private Function ReadLastBatchEnd(stagingRoot As String) As String
    On Error GoTo Fail
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path As String
    path = stagingRoot & "last_batch_end.txt"
    If fso.FileExists(path) Then
        Dim line As String
        line = Trim$(fso.OpenTextFile(path, 1).ReadLine)
        If line Like "####-##-## ##:##" Then ReadLastBatchEnd = line
    End If
    Exit Function
Fail:
    ReadLastBatchEnd = ""
End Function

Private Sub WriteLastBatchEnd(stagingRoot As String, ByVal stamp As Date)
    On Error Resume Next        ' a default that fails to save is not worth a dialog
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(stagingRoot & "last_batch_end.txt", True)
    ts.WriteLine Format(stamp, "yyyy-mm-dd hh:nn")
    ts.Close
End Sub

' mkdir -p. FileSystemObject.CreateFolder makes one level at a time.
Private Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parts() As String
    Dim i As Long
    Dim cur As String
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    parts = Split(path, "\")
    cur = parts(0)                                  ' the drive, "C:"
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Not fso.FolderExists(cur) Then fso.CreateFolder cur
    Next
    Set fso = Nothing
End Sub

Private Sub RemoveFolderIfEmpty(ByVal path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    If fso.GetFolder(path).Files.Count = 0 Then fso.DeleteFolder path
End Sub


