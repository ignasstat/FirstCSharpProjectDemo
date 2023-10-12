Sub LinkFile()

' first need to check the file details are still valid

Dim strFileName As String
Dim strFolder As String
Dim Source As String
Dim strDate As String
Dim CreatedDate As Date
Dim UpdatedDate As Date
Dim ReceivedDate As Date
Dim FileSize As Integer
Dim FileID As Integer
Dim blnProcess As Boolean
Dim strSQL As String
Dim rst As Recordset
'Dim rstJ As Recordset
Set rst = New Recordset
'Set rstJ = New Recordset
Dim selectedRow As Integer
Dim JobNumber As String
Dim IsEmpty As Boolean
Dim FileCheck As String

If Not IsNull(lstFileView.Value) Then

    'Job number from list
    
    lngFileID = lstFileView.Value
    blnProcess = True
    
    'Job number from list
    selectedRow = lstFileView.ListIndex
    selectedRow = selectedRow + 1
    
    JobNumber = lstFileView.Column(1, selectedRow)
    
    If JobNumber = "No Matches" Or JobNumber = "Multiple Matches" Then
        blnProcess = False
        MsgBox ("File has " & JobNumber & " with jobs")
    End If
    
    FileCheck = CheckFileExists(lngFileID, JobNumber)
        
    If blnProcess = True Then
    
            
        ' Check if the file has been processed or selected by another user
        strSQL = "select * from dbo.vw_CT_CheckFile where fileid = " & LTrim(str(lngFileID))
        rst.Open strSQL, db, adOpenForwardOnly, adLockReadOnly
        
        If Not rst.EOF Then
            If Not IsNull(rst("FileUser")) Then
                If LCase(Trim(rst("FileUser"))) = LCase(Trim(strUserName)) Then
                    ' current user so OK
                Else
                    Call MsgBox("File already in use by " & Trim(rst("FileUser")), vbOKOnly)
                    blnProcess = False
                End If
            Else
                ' lock file so nobody else can use it
                strSQL = "insert into  dbo.CT_FileLock (FileID, Source, FileUser, dts) values ("
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
        
        'CTC 039 - Cannot check if file exists directly anymore as user won't have permission to EFT folder, added method CheckFileExists
        If blnProcess Then
            ' need to check if file has changed since last update of the list.
            'If SFTP
            If rst("Source") = "E" Then
                If Not FileCheck = "OK" Then
                    MsgBox ("This file or version of it " & FileCheck & ", Reject it and Refresh Data")
                    Call ClearFileLock(rst("FileID"), "E")
                    blnProcess = False
                End If
            'If Data In
            Else
                If Not FileCheck = "OK" Then
                    MsgBox ("This file or version of it " & FileCheck & ", Reject it and Refresh Data")
                    Call ClearFileLock(rst("FileID"), "I")
                    blnProcess = False
                End If
            End If
        End If
        
        If blnProcess Then
        
                    
            ' Display the file details
            FileID = rst("FileID")
            If rst("Source") = "E" Then
                'lblSource.Caption = "SFTP"
                'Source = "SFTP"
                Source = "E"
            Else
                'lblSource.Caption = "Data In"
                'Source = "Data In"
                Source = "I"
            End If
            
            strFolder = rst("Folder")
            strFileName = rst("FileName")
            CreatedDate = rst("CreatedDate")
            UpdatedDate = rst("UpdatedDate")
            ReceivedDate = rst("SuppliedDate")
            FileSize = rst("FileSize")
            
            'Check if file is empty or has only header
            
            IsEmpty = HeaderOnlyFile(strFolder & strFileName)
            
            If IsEmpty Then
                blnProcess = False
                MsgBox ("This file is empty or contains only header in it")
            End If
            
            If blnProcess Then
            
                If CreatedDate > UpdatedDate Then
                    strDate = Format(CreatedDate, "dd mmm yyyy HH:MM")
                Else
                    strDate = Format(UpdatedDate, "dd mmm yyyy HH:MM")
                End If
                
                
                
                ' create an entry in the Job Run Table
                strSQL = "insert into dbo.CT_JobRun (CT_JobID,fileid,FileSource,RunNo,CreatedBy,createddate,RunStatus, DueByDate) "
                strSQL = strSQL & "Select ct_JobID, "
                strSQL = strSQL & Trim(str(lngFileID)) & ","
                strSQL = strSQL & "'" & Source & "',"
                strSQL = strSQL & "case when LastRun > LastNeptuneRun then LastRun+1 else LastNeptuneRun+1 end, "
                strSQL = strSQL & "'" & strUserName & "',"
                strSQL = strSQL & "getdate(),"
                strSQL = strSQL & "'Logged',"
                strSQL = strSQL & "test('" & strDate & "', j.job_number ) "
                strSQL = strSQL & "from test v "
                strSQL = strSQL & "inner join test j on v.job_number = j.job_number "
                strSQL = strSQL & "where v.job_number = '" & JobNumber & "'"
                
                'MsgBox ("Linking " & Trim(lstJobs.Value))
                MsgBox ("Linking " & strFileName & " to" & JobNumber)
                         
                db.Execute (strSQL)
                
            End If
            
            ' remove the lock on the file
            Call ClearFileLock(Trim(str(lngFileID)), Source)
            
            lstActiveJobs.Requery
            lstFileView.Requery
            pgFileList.Enabled = True
            pgJobList.Enabled = True
            pgJobList.SetFocus
            
            rst.Close
        End If
            
    End If
            

End If


End Sub
