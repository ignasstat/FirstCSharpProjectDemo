        Dim Fileout As TextStream
        Set Fileout = fso.CreateTextFile(fileExistsTriggerFullPath, True, True)
        
        'Set trigger content to full file path
        Fileout.Write strFile
        Fileout.Close
        Set Fileout = Nothing
