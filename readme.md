Function ContainsMultipleLines(FileName As String) As Boolean
    Dim fso As Object
    Dim mf As Object
    Dim LineCount As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists
    If Not fso.FileExists(FileName) Then
        ' File not found
        ContainsMultipleLines = False
    Else
        ' Attempt to open the file
        On Error Resume Next
        Set mf = fso.OpenTextFile(FileName, 1)
        On Error GoTo 0
        
        If Not mf Is Nothing Then
            LineCount = 0
            
            ' Count the lines in the file
            Do While Not mf.AtEndOfStream And LineCount < 2
                mf.ReadLine
                LineCount = LineCount + 1
            Loop
            
            mf.Close
            
            ' Check if the file contains more than one line
            ContainsMultipleLines = (LineCount > 1)
        Else
            ' Unable to open file
            MsgBox "Error Opening File"
            ContainsMultipleLines = False
        End If
    End If
End Function

Function CheckFileExists_Local(myFile As String, ByVal HeaderInFirstRow As Boolean) As String
    Dim fso As Object
    Dim mf As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Nothing to do unless the file exists
    If Not fso.FileExists(myFile) Then
        CheckFileExists_Local = "not found"
    Else
        Set mf = fso.GetFile(myFile)
        
        If mf.Size > 0 Then
            If HeaderInFirstRow Then
                ' If not empty, then need to check how many lines are in it
                If ContainsMultipleLines(myFile) Then
                    CheckFileExists_Local = "OK"
                Else
                    ' File contains only header
                    CheckFileExists_Local = "Is empty"
                End If
            Else
                CheckFileExists_Local = "OK"
            End If
        Else
            CheckFileExists_Local = "Is empty"  ' File is empty
        End If
    End If
End Function
