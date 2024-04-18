Private Function SendFileInternal(JobNumber As String)

'CTC 133 - additional transfer to an internal location
Dim rst As Recordset
Dim strSQL As String
Dim strDeliveryFolder As String

Set rst = New Recordset


strSQL = "select OutputToDOTO, OutputToDSA, OutputToSDQ from dbo.CT_Jobs where Job_Number = '" & JobNumber & "'"
rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly
'strSQL = "select OutputToDOTO, OutputToDSA, OutputToSDQ from dbo.CT_Jobs where Job_Number = '" & Trim(lblJDJobNo.Caption) & "'"
'rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly


If Not rst.EOF Then
    If rst!OutputToDOTO Then
        strDeliveryFolder = GetConfigValue("OutboundDOTO")
        MsgBox strDeliveryFolder
    ElseIf rst!OutputToDSA Then
        strDeliveryFolder = GetConfigValue("OutboundDSA")
        MsgBox strDeliveryFolder
    ElseIf rst!OutputToSDQ Then
        strDeliveryFolder = GetConfigValue("OutboundSDQ")
        MsgBox strDeliveryFolder
        'Call SendFile
    End If
End If

rst.Close

End Function
