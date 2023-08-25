Sub SendEmailViaSMTP()
    Dim objMessage As Object
    Dim objConfig As Object
    Dim objFields As Object
    
    ' Create a message object
    Set objMessage = CreateObject("CDO.Message")
    
    ' Set configuration
    Set objConfig = CreateObject("CDO.Configuration")
    Set objFields = objConfig.Fields
    
    ' Set SMTP server details
    With objFields
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.example.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 ' SMTP port
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' Send using SMTP
        .Update
    End With
    
    ' Configure message properties
    With objMessage
        Set .Configuration = objConfig
        .To = "recipient@example.com"
        .From = "sender@example.com"
        .Subject = "Test Email"
        .TextBody = "This is a test email sent via SMTP using VBA."
        .Send
    End With
    
    ' Clean up
    Set objMessage = Nothing
    Set objConfig = Nothing
    Set objFields = Nothing
End Sub
