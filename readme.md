Function CheckFileExists(lngFileID As Long, JobNumber As String) As String

    Dim rst As Recordset
    Dim strSQL As String
    Dim strFile As String
    Dim HeaderInFirstRow As Boolean

        
    Set rst = New Recordset
    strSQL = "select * from dbo.vw_ExistingFiles where fileid = " & LTrim(str(lngFileID))
    rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly

    If rst.EOF Then
        MsgBox ("file no longer in view (some one else may have processed it)")
        CheckFileExists = "not in view"
        rst.Close
    Else
        strFile = rst("Folder") & rst("Filename")
        rst.Close
        
        'CTC 116 - variable is being used in PS script CalltraceFileExists_ToRelease.ps1. If No header in first row then just checking if a file is empty
        
        HeaderInFirstRow = GetHeaderInFirstRow(JobNumber)
            
        'If a source file is on the local drive
        If InStr(1, strFile, "\\cig.local\Data\Marketing Solutions Departments\Production\", vbTextCompare) > 0 Then
            CheckFileExists = CheckFileExists_Local(strFile, HeaderInFirstRow)
        Else
            CheckFileExists = CheckFileExists_EFT(lngFileID, JobNumber, HeaderInFirstRow, strFile)
        End If
    End If

End Function


Function GetHeaderInFirstRow(JobNumber As String) As Boolean

Dim strSQL As String
Dim rst As Recordset

'Default value is False
GetHeaderInFirstRow = False

strSQL = "SELECT HeadersInFirstRow FROM neptunefileimporter.[fileimporter].[Jobs] "
strSQL = strSQL & "WHERE JobId IN (SELECT Max(JobId) FROM neptunefileimporter.[fileimporter].[Jobs] WHERE DestinationTable = '" & JobNumber & "')"

On Error Resume Next
Set rst = New Recordset
rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly
On Error GoTo 0

If Not rst.EOF Then
    GetHeaderInFirstRow = rst("HeadersInFirstRow")
Else
    MsgBox "Error retrieving HeadersInFirstRow value from neptunefileimporter.[fileimporter].[Jobs]"
End If

rst.Close

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

Function ContainsMultipleLines(FileName As String) As Boolean

Dim fso As Object
Dim mf As Object
Dim LineCount As Long

Set fso = CreateObject("Scripting.FileSystemObject")

ContainsMultipleLines = False

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
    
    If LineCount > 1 Then
        ContainsMultipleLines = True
    Else
        ContainsMultipleLines = False
    End If
    
End If

End Function

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


Function CheckFileExists_EFT(lngFileID As Long, JobNumber As String, ByVal HeaderInFirstRow As Boolean, strFile As String) As String

    Dim fileExistsTriggerFullPath As String
    Dim foundFolder As String
    Dim fileFoundPath As String
    Dim foundNotEmpty As String
    Dim fileFoundNotEmptyPath As String
    Dim foundEmpty As String
    Dim fileFoundEmptyPath As String
    Dim fileNotFoundPath As String
    Dim notFoundFolder As String
    Dim fileExistsTriggerToProcessFolder As String
    
    Dim Fileout As TextStream
            
    'Setup trigger folder paths and paths to the triggers
    fileExistsTriggerToProcessFolder = FileExistsTriggerFolder & "ToProcess\"
    fileExistsTriggerFullPath = fileExistsTriggerToProcessFolder & lngFileID & ".trg"
    
    foundFolder = FileExistsTriggerFolder & "Found\"
    foundNotEmpty = foundFolder & "NotEmpty\"
    foundEmpty = foundFolder & "Empty\"
    notFoundFolder = FileExistsTriggerFolder & "NotFound\"
    
    fileFoundNotEmptyPath = foundNotEmpty & lngFileID & ".trg"
    fileFoundEmptyPath = foundEmpty & lngFileID & ".trg"
    fileNotFoundPath = notFoundFolder & lngFileID & ".trg"
    'fileDonePath
    
    
    'Before checking the trigger file, make sure to delete it if it was created before

    Call DeleteFile(fileFoundNotEmptyPath)
    Call DeleteFile(fileFoundEmptyPath)
    Call DeleteFile(fileNotFoundPath)

    
    'EFT 769 - Adding additional check if a file is already exist.

    Call DeleteFile(fileExistsTriggerFullPath)

    'Create a trigger file called by fileID
    
    Set Fileout = fso.CreateTextFile(fileExistsTriggerFullPath, True, True)
    
    Dim strHeaderInFirstRow As String
    strHeaderInFirstRow = CStr(HeaderInFirstRow)
    
    'Set trigger content to full file path (first row)
    Fileout.Write strFile
    Fileout.Write vbCrLf
    'Adding new line JobNumber (second row)
    Fileout.Write JobNumber
    Fileout.Write vbCrLf
    'Adding new line HeaderInFirstRow (third row)
    Fileout.Write HeaderInFirstRow
    
    Fileout.Close
    Set Fileout = Nothing
    
    'Loop until trigger file appears in one of the three folders
    While Not (fso.FileExists(fileFoundNotEmptyPath) Or (fso.FileExists(fileFoundEmptyPath)) Or (fso.FileExists(fileNotFoundPath)))
    Wend
    
    'Check if trigger file was moved to notfound,FoundNotEmpty and FoundEmpty subfolders
    If fso.FileExists(fileNotFoundPath) Then
        CheckFileExists_EFT = "not found"
        Call DeleteFile(fileNotFoundPath)
    ElseIf fso.FileExists(fileFoundNotEmptyPath) Then
        CheckFileExists_EFT = "OK"
        Call DeleteFile(fileFoundNotEmptyPath)
    ElseIf fso.FileExists(fileFoundEmptyPath) Then
        CheckFileExists_EFT = "Is empty"
        Call DeleteFile(fileFoundEmptyPath)
    End If


End Function
