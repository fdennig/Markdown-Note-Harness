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
            ExtractEmailCore False, False
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
    ExtractEmailCore True, False
End Sub

Sub ExtractEmailWithAttachments()
    ExtractEmailCore False, True
End Sub

Private Sub ExtractEmailCore(toReply As Boolean, saveAttachments As Boolean)

    Dim vaultPathToSaveFileTo As String
    Dim emailFileNameStartChr As String
    Dim emailTypeLink As String
    Dim personNameStartChar As String
    Dim attachmentPath As String
    config vaultPathToSaveFileTo, personNameStartChar, emailFileNameStartChr, emailTypeLink, attachmentPath

    '================================================'
    ' Save as plain text
    Const OLTXT = 0
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
    Dim temporarySubjectLineString As String
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

        ' Yank the mail items subject line to `temporarySubjectLineString`
        temporarySubjectLineString = oMail.Subject
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
        resultString = resultString & "class: email" & vbCrLf
    resultString = resultString & "subject: """ & temporarySubjectLineString & """" & vbCrLf
        resultString = resultString & "date: " & Format(oMail.ReceivedTime, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "from: """ & sender & """" & vbCrLf
        ' Add sender email address if available (placed once here)
        Dim senderEmail As String
        senderEmail = GetSenderSmtp(oMail)
        If Len(senderEmail) > 0 Then
            resultString = resultString & "from_email: """ & senderEmail & """" & vbCrLf
        End If
        ' Convert recipients to YAML list
        resultString = resultString & "to:" & vbCrLf
        For Each recip In recips
            resultString = resultString & "  - """ & formatName(recip.Name, personNameStartChar) & """" & vbCrLf
        Next
        resultString = resultString & "tags:" & vbCrLf
        ' Attachments list inside front matter
        Dim att As Attachment, attName As String
        Dim attPrefix As String
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

    ' End YAML block then blank line
    resultString = resultString & "---" & vbCrLf & vbCrLf
    ' Add a Markdown task before the email body
    ' Add original email content
    resultString = resultString & ConvertLinksToMarkdown(oMail.Body)


        ' Make a dummy email to hold the details we're saving
        ' This way we dont get junk in the message header when saving
        Dim outputItem As MailItem
            Set outputItem = Application.CreateItem(olMailItem)
        outputItem.Body = resultString

        ' Build improved filename for fuzzy search:
        ' Format: <prefix><yyyy-MM-dd>_<HHNN> @Sender - Subject.md
        ' Ensures: single underscore between date/time, single @, normalized sender spacing.
        Dim dateToken As String: dateToken = Format(dtDate, "yyyy-MM-dd", vbUseSystemDayOfWeek, vbUseSystem)
        Dim timeToken As String: timeToken = Format(dtDate, "HHNN", vbUseSystemDayOfWeek, vbUseSystem)
        Dim senderToken As String: senderToken = sender
        On Error Resume Next
        ReplaceCharsForFileName senderToken, " "
        On Error GoTo EndClean
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

    ' Save the result
    SaveAsUTF8 vaultPathToSaveFileTo & fileName, resultString

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
    ShellExecute 0, "open", "subl", """" & vaultPathToSaveFileTo & fileName & """", vbNullString, 1


    Next
EndClean:
    Set obj = Nothing
    Set oMail = Nothing
    Set outputItem = Nothing
End Sub

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

        ' Clean subject for filename
        Dim temporarySubjectLineString As String
        temporarySubjectLineString = oAppt.Subject
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
        resultString = resultString & "subject: """ & temporarySubjectLineString & """" & vbCrLf
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

        resultString = resultString & "---" & vbCrLf & vbCrLf

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

        ' Open in Sublime Text
        ShellExecute 0, "open", "subl", """" & vaultPathToSaveFileTo & fileName & """", vbNullString, 1

    Next
EndCleanMeeting:
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
            ' PR_SMTP_ADDRESS
            addr = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        End If
    End If
    GetSenderSmtp = LCase$(addr)
    Exit Function
Fail:
    GetSenderSmtp = ""
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


