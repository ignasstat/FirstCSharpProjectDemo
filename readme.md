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
