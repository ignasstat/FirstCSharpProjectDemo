  Dim strFileName As String
    Dim Source As String
    Dim strDate As String
    Dim CreatedDate As Date
    Dim UpdatedDate As Date
    Dim FileSize As Integer
    Dim blnProcess As Boolean
    Dim strSQL As String
    Dim rst As Recordset
    
    Set rst = New Recordset
    
    ' Set rstJ = New Recordset
    Dim selectedRow As Integer
    Dim JobNumber As String
    Dim IsEmpty As Boolean
    Dim FileCheck As String
    Dim CanLaunchView As String
    
    
    If Not IsNull(lstFileView.Value) Then
        ' Job number from list
        lngFileID = lstFileView.Value
        blnProcess = True
        
        ' Job number from list
        selectedRow = lstFileView.ListIndex
        selectedRow = selectedRow + 1
        JobNumber = lstFileView.Column(1, selectedRow)
        
        ' Can Launch
        CanLaunchView = lstFileView.Column(5, selectedRow)
        
        If JobNumber = "No Matches" Or JobNumber = "Multiple Matches" Then
            blnProcess = False
            MsgBox ("File has " & JobNumber & " with jobs")
        ElseIf JobNumber Like "*NA" Then
            blnProcess = False
            MsgBox ("Cannot process as " & JobNumber & " is inactive")
        End If
        
        If chkLaunchJob.Value = "True" And CanLaunchView = "False" Then
            blnProcess = False
            MsgBox "Cannot launch " & JobNumber & ". To Setup Job please deselect launching option." & vbNewLine & "Alternatively investigate an issue."
        End If
        
        ' CTC116 Additional parameter JobNumber, also function returns string instead of boolean
        FileCheck = CheckFileExists(lngFileID, JobNumber)
        
        'If a file is not present in dbo.vw_ExistingFiles
        If FileCheck = "not in view" Then
            blnProcess = False
        End If
        'If a trigger file failed to be created
        If FileCheck = "trigger notfound" Then
            blnProcess = False
        End If
        
        If blnProcess = True Then
            ' Check if the file has been processed or selected by another user
            strSQL = "select * from dbo.vw_CT_CheckFile where fileid = " & LTrim(str(lngFileID))
            rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly
            
            If Not rst.EOF Then
                If Not IsNull(rst("FileUser")) Then
                    If LCase(Trim(rst("FileUser"))) = LCase(Trim(strUserName)) Then
                        ' Current user, so OK
                    Else
                        Call MsgBox("File already in use by " & Trim(rst("FileUser")), vbOKOnly)
                        blnProcess = False
                    End If
                Else
                    ' Lock the file so nobody else can use it
                    strSQL = "insert into dbo.CT_FileLock (FileID, Source, FileUser, dts) values ("
                    strSQL = strSQL & rst("FileID") & ","
                    strSQL = strSQL & "'" & rst("Source") & "',"
                    strSQL = strSQL & "'" & strUserName & "',"
                    strSQL = strSQL & "getdate())"
                    db.Execute (strSQL)
                End If
            Else
                blnProcess = False
                MsgBox ("File no longer available to process, Reject it and Refresh Data")
                lstFileView.Requery
                lstRejected.Requery
            End If
            
            ' CTC 039 - Cannot check if the file exists directly anymore as the user won't have permission to the EFT folder, added method CheckFileExists
            If blnProcess Then
                ' Need to check if the file has changed since the last update of the list.
                ' If SFTP
                If rst("Source") = "E" Then
                    If Not FileCheck = "OK" Then
                        MsgBox ("This file or version of it " & FileCheck & ", Reject it and Refresh Data")
                        Call ClearFileLock(rst("FileID"), "E")
                        blnProcess = False
                    End If
                ' If Data In
                Else
                    If Not FileCheck = "OK" Then
                        MsgBox ("This file or version of it " & FileCheck & ", Reject it and Refresh Data")
                        Call ClearFileLock(rst("FileID"), "I")
                        blnProcess = False
                    End If
                End If
            End If
            
            If blnProcess Then
                If rst("Source") = "E" Then
                    ' lblSource.Caption = "SFTP"
                    ' Source = "SFTP"
                    Source = "E"
                Else
                    ' lblSource.Caption = "Data In"
                    ' Source = "Data In"
                    Source = "I"
                End If
                
                CreatedDate = rst("CreatedDate")
                UpdatedDate = rst("UpdatedDate")
            
                If blnProcess Then
                    If CreatedDate > UpdatedDate Then
                        strDate = Format(CreatedDate, "dd mmm yyyy HH:MM")
                    Else
                        strDate = Format(UpdatedDate, "dd mmm yyyy HH:MM")
                    End If
                    
                    ' Create an entry in the Job Run Table
                    strSQL = "insert into dbo.CT_JobRun (CT_JobID,fileid,FileSource,RunNo,CreatedBy,createddate,RunStatus, DueByDate) "
                    strSQL = strSQL & "Select ct_JobID, "
                    strSQL = strSQL & Trim(str(lngFileID)) & ","
                    strSQL = strSQL & "'" & Source & "',"
                    strSQL = strSQL & "case when LastRun > LastNeptuneRun then LastRun+1 else LastNeptuneRun+1 end, "
                    strSQL = strSQL & "'" & strUserName & "',"
                    strSQL = strSQL & "getdate(),"
                    strSQL = strSQL & "'Logged',"
                    strSQL = strSQL & "dbo.fn_CallTraceDueDate_New('" & strDate & "', j.job_number ) "
                    strSQL = strSQL & "from dbo.vw_CallTraceJobList v "
                    strSQL = strSQL & "inner join dbo.ct_jobs j on v.job_number = j.job_number "
                    strSQL = strSQL & "where v.job_number = '" & JobNumber & "'"
                    
                    ' MsgBox ("Linking " & Trim(lstJobs.Value))
                    strFileName = rst("FileName")
                    MsgBox ("Linking " & strFileName & " to " & JobNumber)
                    
                    db.Execute (strSQL)
                End If
                
                rst.Close
                ' Remove the lock on the file
                Call ClearFileLock(Trim(str(lngFileID)), Source)
                         
                'SETUP JOB part
                
                Dim RunID As String
                Dim JobRunStatus As String
                Dim RunNumber As Integer
                Dim FileFolder As String
                Dim FileName As String
                Dim JobClient As String
                Dim JobType As String
                
                strSQL = "select CT_RunID from dbo.CT_JobRun where FileID = " & lngFileID
                rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly
                             
                RunID = rst("CT_RunID")
                
                rst.Close
                     
                strSQL = "select * from dbo.vw_CallTrace_JobDetail where runid = " & RunID
                rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly
                
           
                JobRunStatus = rst("Status")
                RunNumber = rst("RunNo")
                FileFolder = rst("folder")
                FileName = rst("FileName")
                JobClient = rst("Client")
                RunID = rst("RunID")
                JobType = rst("JobType")
                
                rst.Close

                Call SetupJob(JobRunStatus, JobNumber, RunNumber, FileFolder, FileName, JobClient, RunID)


                -----------------

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
    Dim attemptCount As Integer
    Dim maxAttempts As Integer
    
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
    'CTC 127
    attemptCount = 0
    maxAttempts = 3
    
    Do While attemptCount < maxAttempts
       If (fso.FileExists(fileFoundNotEmptyPath) Or fso.FileExists(fileFoundEmptyPath) Or fso.FileExists(fileNotFoundPath)) Then
           ' File found, break out of the loop
           Exit Do
       End If
       attemptCount = attemptCount + 1
       Sleep (3)
    Loop
    
    If attemptCount = maxAttempts Then
        MsgBox "Trigger file not found after multiple attempts."
        CheckFileExists_EFT = "trigger notfound"
        Exit Function
    End If
    
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


------------

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
'CTC 121
Function DeleteFile(filePath As String)
    Dim fso As Object
    Dim attemptCount As Integer
    Dim maxAttempts As Integer
    Dim success As Boolean
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists before attempting to delete
    If Not fso.FileExists(filePath) Then
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

-----------

Public Sub ClearFileLock(FileID As String, FileSource As String)
' remove the lock on a file if it exists

Dim SQL As String

SQL = "execute [dbo].[up_CT_ClearFileLock] " & FileID & ", '" & FileSource & "'"
db.Execute (SQL)

End Sub

-------------

                
