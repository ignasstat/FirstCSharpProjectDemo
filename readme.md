Function DeleteFile(filePath As String)
    Dim fso As Object
    Dim attemptCount As Integer
    Dim maxAttempts As Integer
    Dim success As Boolean
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists before attempting to delete
    If Not fso.FileExists(filePath) Then
        MsgBox "File '" & filePath & "' does not exist."
        Exit Function
    End If
    
    attemptCount = 0
    maxAttempts = 6
    
    Do While attemptCount < maxAttempts
        ' Attempt to delete the file
        On Error Resume Next
        fso.DeleteFile filePath
        On Error GoTo 0
        
        ' Check if the file still exists after deletion attempt
        If Not fso.FileExists(filePath) Then
            success = True
            Exit Do
        End If
        
        ' Wait for 10 seconds before the next deletion attempt
        Sleep (10)
        
        attemptCount = attemptCount + 1
    Loop
    
    ' Message if the file could not be deleted after multiple attempts
    If Not success Then
        MsgBox "Trigger file '" & filePath & "' could not be deleted after multiple attempts."
    End If
      
End Function
