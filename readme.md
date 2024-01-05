Function GetHeadersInFirstRow(JobNumber As String) As Variant
    Dim strSQL As String
    Dim rst As Recordset
    
    ' Assuming db is a valid database object, adOpenForwardOnly and adLockReadOnly are constants defined elsewhere
    
    ' Initialize the return value
    GetHeadersInFirstRow = False
    
    ' Construct SQL query
    strSQL = "SELECT HeadersInFirstRow FROM test "
    strSQL = strSQL & "WHERE JobId IN (SELECT Max(JobId) FROM test WHERE DestinationTable = '" & JobNumber & "')"
    
    ' Execute the query
    On Error Resume Next
    Set rst = New Recordset
    rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly
    On Error GoTo 0
    
    ' Check if the recordset is not empty and no error occurred
    If Not rst.EOF And rst.State = adStateOpen Then
        GetHeadersInFirstRow = rst("HeadersInFirstRow")
    Else
        ' Display error message if needed
        MsgBox "Error retrieving HeadersInFirstRow value from test"
    End If
    
    ' Close the recordset
    If Not rst Is Nothing Then
        rst.Close
    End If
End Function
