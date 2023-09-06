Function sendEmail(strTo As String, strSubject As String, strTextBody As String, strCC As String, Optional attachements As Variant)

Dim objMessage As Object
Dim objConfig As Object
Dim objFields As Object

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
    ' Attach files from the provided array
    Dim i As Integer
    If attachements <> Empty Then
        For i = LBound(attachements) To UBound(attachements)
            If Len(attachements(i)) > 0 Then
                .AddAttachment attachements(i)
            End If
        Next i
    End If
    
    .send
End With

'Send the email


Set objMessage = Nothing
Set objConfig = Nothing
Set objFields = Nothing

End Function
