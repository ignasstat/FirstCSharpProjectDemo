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


-------------------


Private Sub Command461_Click()

Dim strAttachFolder As String
Dim myCount As Integer
Dim attachements() As String
Dim strSentFile As String
Dim strBody As String

strAttachFolder = "\\cig.local\data\AppData\SFTP\Data\Usr\DataBureau\Configuration\Scripts\Test\CallTrace Console\CTC100\Attachments\"

    strBody = "<html>"
    strBody = strBody & "<body style = ""font-family:arial,helvetica,sans-serif;font-size:11pt;"">"
    strBody = strBody & "<p>Hello</p>"
    strBody = strBody & "<p>The data for this job has been output as follows</p>"
 
    strBody = strBody & "<table style = ""font-family:arial,helvetica,sans-serif;font-size:11pt;margin:0px;padding:0px;border-collapse:collapse;"">"
    strBody = strBody & "   <tr>"
    strBody = strBody & "       <th style = ""margin:0px;padding:5px;background-color:#00A6CA;color:#ffffff;text-align:left;"">&nbsp;</th>"
    strBody = strBody & "       <th style = ""margin:0px;padding:5px;background-color:#00A6CA;color:#ffffff;text-align:left;"">File</th>"
    strBody = strBody & "       <th style = ""margin:0px;padding:5px;background-color:#00A6CA;color:#ffffff;text-align:left;"">Date</td>"
    strBody = strBody & "   </tr>"
    strBody = strBody & "   <tr>"
    strBody = strBody & "       <td style = ""font-weight:bold;font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;"">Input File</td>"
    strBody = strBody & "       <td style = ""font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;""> Hello </td>"
    strBody = strBody & "       <td style = ""font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;""> Den </td>"
    strBody = strBody & "   </tr>"
    strBody = strBody & "   <tr>"
    strBody = strBody & "       <td style = ""font-weight:bold;font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;vertical-align:top;"">Output Files</td>"
    strBody = strBody & "       <td style = ""font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;"">"
                

        
    ' attach required files
    strSentFile = Dir(strAttachFolder & "*.*")
    While Len(strSentFile) > 0
        'OutMail.Attachments.Add (strAttachFolder & strSentFile)
        myCount = myCount + 1
        ReDim Preserve attachements(1 To myCount)
        attachements(myCount) = strAttachFolder & strSentFile
        strSentFile = Dir

        If InStr(LCase(strSentFile), "_Error_Report.") Then
            strBody = strBody & "<p>Error Report Attached</p>"
        ElseIf InStr(LCase(strSentFile), "_Duplicates_Report.") Then
            strBody = strBody & "<p>Error Duplicate Report Attached</p>"
        ElseIf InStr(LCase(strSentFile), "_Report.") Then
            strBody = strBody & "<p>Report Attached</p>"
        End If
        
       ' strSentFile = Dir
    Wend
    
    ' #CTR053 - 24/09/2018 - output notification text changed
    strBody = strBody & "<p>Questions, queries, issues and additional requirements or changes in relation to this job should be directed to the Data Bureau team using Test Track Pro by filling in the ""Complaint Tracking"" section.</p>"
    strBody = strBody & "<p>Regards</p>"
    strBody = strBody & "</body>"
    strBody = strBody & "</html>"

    Dim strTo As String
    Dim strFrom As String
    Dim strSubject As String
    Dim strTextBody As String
    Dim strCC As String

    strTo = "ignas.statkevicius@transunion.com"
    strFrom = "ignas.statkevicius@transunion.com"
    strSubject = "Test Email"
    strTextBody = strBody



    Call sentEmail(strTo, strFrom, strSubject, strTextBody, attachements)

        
End Sub

Function sentEmail(strTo As String, strFrom As String, strSubject As String, strTextBody As String, attachements() As String, strCC As String)

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
    .From = strFrom
    .Subject = strSubject
    .TextBody = strTextBody
    ' Attach files from the provided array
    Dim i As Integer
    For i = LBound(attachements) To UBound(attachements)
        If Len(attachements(i)) > 0 Then
            .AddAttachment Attach
        End If
    Next i
End With

'Send the email
.send

Set objMessage = Nothing
Set objConfig = Nothing
Set objFields = Nothing

End Function


<html><body style = "font-family:arial,helvetica,sans-serif;font-size:11pt;"><p>Hello</p><p>The data for this job has been output as follows</p><table style = "font-family:arial,helvetica,sans-serif;font-size:11pt;margin:0px;padding:0px;border-collapse:collapse;">   <tr>       <th style = "margin:0px;padding:5px;background-color:#00A6CA;color:#ffffff;text-align:left;">&nbsp;</th>       <th style = "margin:0px;padding:5px;background-color:#00A6CA;color:#ffffff;text-align:left;">File</th>       <th style = "margin:0px;padding:5px;background-color:#00A6CA;color:#ffffff;text-align:left;">Date</td>   </tr>   <tr>       <td style = "font-weight:bold;font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;">Input File</td>       <td style = "font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;"> Hello </td>       <td style = "font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;"> Den </td>   </tr>   <tr>       <td style = "font-weight:bold;font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;vertical-align:top;">Output Files</td>       <td style = "font-size:11pt;margin:0px;padding:5px;border-bottom-style:solid;border-width:1px;border-color:#999999;"><p>Questions, queries, issues and additional requirements or changes in relation to this job should be directed to the Data Bureau team using Test Track Pro by filling in the "Complaint Tracking" section.</p><p>Regards</p></body></html>
