Function sendEmail(strTo As String, strSubject As String, strTextBody As String, strCC As String, Optional attachments As Variant)

    Dim objMessage As Object
    Dim objConfig As Object
    Dim objFields As Object
    Dim i As Integer
    
    On Error Resume Next ' Turn on error handling

    Set objMessage = CreateObject("CDO.Message")

    ' Set configuration
    Set objConfig = CreateObject("CDO.Configuration")
    Set objFields = objConfig.Fields

    ' Set SMTP server details
    With objFields
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtprelay.cig.local"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 ' SMTP port
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' Send using SMTP
        .Update
    End With

    ' Configure message properties
    With objMessage
        Set .Configuration = objConfig
        .To = strTo
        .CC = strCC
        If EmailAs <> "" Then
            .From = EmailAs
        End If
        .Subject = strSubject
        .HTMLBody = strTextBody
        
        ' Attach files from the provided array (if it's an array)
        If IsArray(attachments) Then
            For i = LBound(attachments) To UBound(attachments)
                If Len(attachments(i)) > 0 Then
                    .AddAttachment attachments(i)
                End If
            Next i
        End If

        .Send
    End With

    ' Error handling
    If Err.Number <> 0 Then
        ' Handle the error here (you can display a message or log it)
        Debug.Print "Error sending email: " & Err.Description
    End If

    On Error GoTo 0 ' Reset error handling
    
    ' Clean up
    Set objMessage = Nothing
    Set objConfig = Nothing
    Set objFields = Nothing

End Function
