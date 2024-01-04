strSQL = "SELECT HeadersInFirstRow FROM test "
        strSQL = strSQL & "WHERE JobId IN (SELECT Max(JobId) FROM test WHERE DestinationTable =  '" & JobNumber & "'" & ")"
        rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly
        'When
        If rst.EOF Then
            MsgBox ("Cannot get HeadersInFirstRow a value form test")
            rst.Close
        Else
            HeaderInFirstRow = rst("HeadersInFirstRow")
            rst.Close
