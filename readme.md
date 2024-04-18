Private Function SendFileInternal(JobNumber As String)
    ' CTC 133 - additional transfer to an internal location
    Dim rst As Recordset
    Dim strSQL As String
    Dim strDeliveryFolder As String
    Dim db As Database ' Ensure that db is declared and set appropriately

    ' Initialize the database connection
    Set db = CurrentDb()

    ' SQL query to select columns based on the job number
    strSQL = "SELECT OutputToDOTO, OutputToDSA, OutputToSDQ FROM dbo.CT_Jobs WHERE Job_Number = '" & JobNumber & "'"

    On Error GoTo ErrorHandler ' Start error handling

    Set rst = db.OpenRecordset(strSQL, dbOpenForwardOnly, adLockReadOnly)

    If Not rst.EOF Then
        ' Check each condition separately
        If rst!OutputToDOTO Then
            strDeliveryFolder = GetConfigValue("OutboundDOTO")
            MsgBox strDeliveryFolder
            ' Potentially call SendFile or other functionality
        End If
        If rst!OutputToDSA Then
            strDeliveryFolder = GetConfigValue("OutboundDSA")
            MsgBox strDeliveryFolder
            ' Potentially call SendFile or other functionality
        End If
        If rst!OutputToSDQ Then
            strDeliveryFolder = GetConfigValue("OutboundSDQ")
            MsgBox strDeliveryFolder
            ' Potentially call SendFile or other functionality
        End If
    End If

Cleanup:
    ' Clean up resources
    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
        Set rst = Nothing
    End If
    Set db = Nothing
    Exit Function

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume Cleanup
End Function
