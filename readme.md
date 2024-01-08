Function CheckFileExists_Local(myFile As String, ByVal HeaderInFirstRow As Boolean) As String
    Dim fso As Object
    Dim mf As Object
    Dim Dummy As String
    Dim LineCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Nothing to do unless the file exists
    If Not fso.FileExists(myFile) Then
        CheckFileExists_Local = "not found"
    Else
        Set mf = fso.GetFile(myFile)
        
        If mf.Size > 0 Then
            If HeaderInFirstRow Then
                ' If not empty, then need to check how many lines are in it
                On Error Resume Next
                Set mf = fso.OpenTextFile(myFile, 1)
                On Error GoTo 0
                
                If Not mf Is Nothing Then
                    LineCount = 0
                    
                    Do While Not mf.AtEndOfStream And LineCount < 2
                        Dummy = mf.ReadLine
                        LineCount = LineCount + 1
                    Loop
                    
                    mf.Close
                    
                    If LineCount > 1 Then
                        CheckFileExists_Local = "OK"
                    Else
                    'File contains only header
                        CheckFileExists_Local = "Is empty"
                    End If
                Else
                    ' Unable to open file
                    MsgBox "Error Opening File"
                    CheckFileExists_Local = "not found"
                End If
            Else
                CheckFileExists_Local = "OK"
            End If
        Else
            CheckFileExists_Local = "Is empty"  ' File is empty
        End If
    End If
End Function
