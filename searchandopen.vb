Sub SearchAndOpenEmailInReplyAllMode()

    Dim myNamespace As Outlook.NameSpace
    Dim myFolder As Outlook.folder
    Dim mySearch As String
    Dim mySearchResult As Outlook.items
    Dim searchCriteria As String
    Dim myMail As Outlook.MailItem
    Dim replyMail As Outlook.MailItem
    Dim subjectWithoutRE As String
    Dim rec As Outlook.Recipient
    Dim bccRecipients As String

    Set myNamespace = Application.GetNamespace("MAPI")
    mySearch = InputBox("Enter the subject line or change number of the email you want to search for:", "Email Search")

    ' Get the "Citywide Service Desk" mailbox
    Set myFolder = myNamespace.Folders("Citywide Service Desk").Folders("Sent Items")

    ' Determine if the search is for a change number or a full subject line
    If mySearch Like "CHG*" Then
        ' Search for change number within the subject line
        searchCriteria = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " LIKE '%" & mySearch & "%' AND " & _
                         Chr(34) & "urn:schemas:httpmail:date" & Chr(34) & " >= '" & Format$(Date - 30, "yyyy-mm-dd") & "' AND " & _
                         Chr(34) & "urn:schemas:httpmail:date" & Chr(34) & " <= '" & Format$(Date + 1, "yyyy-mm-dd") & "'"
    Else
        ' Search for exact subject line
        searchCriteria = "[Subject] = '" & mySearch & "' AND [SentOn] >= '" & Format(Date - 30, "yyyy-mm-dd") & "' AND [SentOn] <= '" & Format(Date + 1, "yyyy-mm-dd") & "'"
    End If

    Set mySearchResult = myFolder.items.Restrict(searchCriteria)
    mySearchResult.Sort "[SentOn]", True ' Sort by SentOn descending

    If mySearchResult.Count > 0 Then
        Set myMail = mySearchResult.GetFirst
        Set replyMail = myMail.ReplyAll

        ' Store BCC recipients from the original mail
        bccRecipients = myMail.BCC

        ' Remove all recipients from To and CC fields
        For i = replyMail.Recipients.Count To 1 Step -1
            Set rec = replyMail.Recipients.Item(i)
            If rec.Type = olTo Or rec.Type = olCC Then
                rec.Delete
            End If
        Next i

        ' Add BCC recipients to the reply mail
        If bccRecipients <> "" Then
            Dim bccArr As Variant
            bccArr = Split(bccRecipients, ";")
            For i = LBound(bccArr) To UBound(bccArr)
                replyMail.BCC = replyMail.BCC & ";" & bccArr(i)
            Next i
        End If

        ' Remove "RE:" prefix from subject if present (not case-sensitive)
        subjectWithoutRE = Trim(myMail.Subject)
        If InStr(1, UCase(subjectWithoutRE), "RE:", vbTextCompare) = 1 Then
            subjectWithoutRE = Trim(Mid(subjectWithoutRE, Len("RE:") + 1))
        End If
        replyMail.Subject = subjectWithoutRE
        
        ' Set the From field to the specified address
        replyMail.SentOnBehalfOfName = "youremail@example.com"

        replyMail.Display
    Else
        MsgBox "No email found with the subject line or change number '" & mySearch & "' in the mailbox."
    End If

    ' Clean up
    Set myNamespace = Nothing
    Set myFolder = Nothing
    Set mySearchResult = Nothing
    Set myMail = Nothing
    Set replyMail = Nothing

End Sub