Sub CreateDecryptionNode(ByVal GUIDString As String, ByRef ActionListNames As String, ByVal SaveLocation As String, ByRef nodeStep As Integer, ByVal FileName As String, ByVal SourceFolder As String, ByVal PGPFolder As String, ByVal PGP_Name As String, ByVal JobNumber As String, ByVal JobClient As String)
    
    Dim XML_Action As IXMLDOMElement
    Dim XML_Child  As IXMLDOMElement
    
    Dim XML_Doc             As DOMDocument
    Dim XML_RootElement     As IXMLDOMElement
    Dim XML_Attribute       As IXMLDOMAttribute
    Dim XML_Element         As IXMLDOMElement
    Dim ActionXMLPath       As String
    
    'Setup the xml
    Set XML_Doc = New DOMDocument
    XML_Doc.async = False
    XML_Doc.validateOnParse = False
    
    'Create temporary folder for decryption
    '----------------------------------------------------------------------------------------------------------------------------------------
    Set XML_Action = XML_Doc.createElement("EFT_Action")
    XML_Doc.appendChild XML_Action
    
    Call XML_AddNode(XML_Doc, XML_Action, "ProcessPriority", "50")
    Call XML_AddNode(XML_Doc, XML_Action, "GUID", GUIDString)
    Call XML_AddNode(XML_Doc, XML_Action, "StepSequence", CStr(nodeStep))
    Call XML_AddNode(XML_Doc, XML_Action, "CreationDTS", Format(Now, "yyyy-MM-dd hh:mm:ss"))
    Call XML_AddNode(XML_Doc, XML_Action, "RetriesRemaining", "2")
    Call XML_AddNode(XML_Doc, XML_Action, "LastAttemptDTS", "")
    Call XML_AddNode(XML_Doc, XML_Action, "LastResult", "")
    Call XML_AddNode(XML_Doc, XML_Action, "FailureNotificationEmail", "")
    
    'Populate FailureNotificationEmail child notes with email info
    '----------------------------------
    Set XML_Child = XML_Action.LastChild
    Call XML_AddNode(XML_Doc, XML_Child, "To", "DataOperationsEFT-CD@transunion.co.uk")
    Call XML_AddNode(XML_Doc, XML_Child, "CC", "DataBureau@transunion.co.uk")
    Call XML_AddNode(XML_Doc, XML_Child, "BCC", "")
    Call XML_AddNode(XML_Doc, XML_Child, "Subject", "DTP - " & JobClient & " - " & JobNumber & " - Folder Creation - Failed")
    Call XML_AddNode(XML_Doc, XML_Child, "Message", "The DTP has failed at creating folder while decrypting " & PGPFolder & vbCrLf)
    '----------------------------------
    
    Call XML_AddNode(XML_Doc, XML_Action, "ActionType", "CreateFolder")
    Call XML_AddNode(XML_Doc, XML_Action, "SourceFolder", "")
    Call XML_AddNode(XML_Doc, XML_Action, "SourceFile", "")
    Call XML_AddNode(XML_Doc, XML_Action, "DestinationFolder", PGPFolder)
    Call XML_AddNode(XML_Doc, XML_Action, "DestinationFile", "")
    
    'Need to create a name for Action.xml
    'Pad the nodeStep with leading zero if neccessary
    ActionXMLPath = SaveLocation & "50_" & GUIDString & "_" & Format(nodeStep, "#00") & ".xml"
    
    XML_Doc.Save (ActionXMLPath)
    ActionListNames = ActionListNames & ActionXMLPath & ";"
    
    nodeStep = nodeStep + 1
    '----------------------------------------------------------------------------------------------------------------------------------------
    
    'Copy to temporary decryption folder
    '----------------------------------------------------------------------------------------------------------------------------------------
    'Setup the xml
    Set XML_Doc = New DOMDocument
    XML_Doc.async = False
    XML_Doc.validateOnParse = False
    
    Set XML_Action = XML_Doc.createElement("EFT_Action")
    XML_Doc.appendChild XML_Action
    
    Call XML_AddNode(XML_Doc, XML_Action, "ProcessPriority", "50")
    Call XML_AddNode(XML_Doc, XML_Action, "GUID", GUIDString)
    Call XML_AddNode(XML_Doc, XML_Action, "StepSequence", CStr(nodeStep))
    Call XML_AddNode(XML_Doc, XML_Action, "CreationDTS", Format(Now, "yyyy-MM-dd hh:mm:ss"))
    Call XML_AddNode(XML_Doc, XML_Action, "RetriesRemaining", "2")
    Call XML_AddNode(XML_Doc, XML_Action, "LastAttemptDTS", "")
    Call XML_AddNode(XML_Doc, XML_Action, "LastResult", "")
    Call XML_AddNode(XML_Doc, XML_Action, "FailureNotificationEmail", "")
    
    'Populate FailureNotificationEmail child notes with email info
    '----------------------------------
    Set XML_Child = XML_Action.LastChild
    Call XML_AddNode(XML_Doc, XML_Child, "To", "DataOperationsEFT-CD@transunion.co.uk")
    Call XML_AddNode(XML_Doc, XML_Child, "CC", "DataBureau@transunion.co.uk")
    Call XML_AddNode(XML_Doc, XML_Child, "BCC", "")
    Call XML_AddNode(XML_Doc, XML_Child, "Subject", "DTP - " & JobClient & " - " & JobNumber & " - File Transfer - Failed")
    Call XML_AddNode(XML_Doc, XML_Child, "Message", "The DTP has failed before decryption to copy " & FileName & vbCrLf & "From: " & SourceFolder & vbCrLf & "To: " & PGPFolder & vbCrLf)
    '----------------------------------
    
    Call XML_AddNode(XML_Doc, XML_Action, "ActionType", "Copy")
    Call XML_AddNode(XML_Doc, XML_Action, "SourceFolder", SourceFolder)
    Call XML_AddNode(XML_Doc, XML_Action, "SourceFile", FileName)
    Call XML_AddNode(XML_Doc, XML_Action, "DestinationFolder", PGPFolder)
    Call XML_AddNode(XML_Doc, XML_Action, "DestinationFile", "")
    
    'Need to create a name for Action.xml
    'Pad the nodeStep with leading zero if neccessary
    ActionXMLPath = SaveLocation & "50_" & GUIDString & "_" & Format(nodeStep, "#00") & ".xml"
    
    XML_Doc.Save (ActionXMLPath)
    ActionListNames = ActionListNames & ActionXMLPath & ";"
    
    nodeStep = nodeStep + 1
    '----------------------------------------------------------------------------------------------------------------------------------------
    
    'Decryption Action
    '----------------------------------------------------------------------------------------------------------------------------------------
    'Setup the xml
    Set XML_Doc = New DOMDocument
    XML_Doc.async = False
    XML_Doc.validateOnParse = False
    
    Set XML_Action = XML_Doc.createElement("EFT_Action")
    XML_Doc.appendChild XML_Action
    
    Call XML_AddNode(XML_Doc, XML_Action, "ProcessPriority", "50")
    Call XML_AddNode(XML_Doc, XML_Action, "GUID", GUIDString)
    Call XML_AddNode(XML_Doc, XML_Action, "StepSequence", CStr(nodeStep))
    Call XML_AddNode(XML_Doc, XML_Action, "CreationDTS", Format(Now, "yyyy-MM-dd hh:mm:ss"))
    Call XML_AddNode(XML_Doc, XML_Action, "RetriesRemaining", "2")
    Call XML_AddNode(XML_Doc, XML_Action, "LastAttemptDTS", "")
    Call XML_AddNode(XML_Doc, XML_Action, "LastResult", "")
    Call XML_AddNode(XML_Doc, XML_Action, "FailureNotificationEmail", "")
    
    'Populate FailureNotificationEmail child nodes with email info
    '---------------------------------------------
    Set XML_Child = XML_Action.LastChild
    Call XML_AddNode(XML_Doc, XML_Child, "To", "DataOperationsEFT-CD@transunion.co.uk")
    Call XML_AddNode(XML_Doc, XML_Child, "CC", "DataBureau@transunion.co.uk")
    Call XML_AddNode(XML_Doc, XML_Child, "BCC", "")
    Call XML_AddNode(XML_Doc, XML_Child, "Subject", "DTP - " & JobClient & " - " & JobNumber & " - File Decryption - Failed")
    Call XML_AddNode(XML_Doc, XML_Child, "Message", "The DTP has failed at decrypting " & FileName & " in " & PGPFolder & vbCrLf)
    '---------------------------------------------
    
    Call XML_AddNode(XML_Doc, XML_Action, "ActionType", "Decrypt")
    Call XML_AddNode(XML_Doc, XML_Action, "SourceFolder", PGPFolder)
    Call XML_AddNode(XML_Doc, XML_Action, "SourceFile", FileName)
    Call XML_AddNode(XML_Doc, XML_Action, "DestinationFolder", "")
    Call XML_AddNode(XML_Doc, XML_Action, "DestinationFile", "")
    Call XML_AddNode(XML_Doc, XML_Action, "EncryptionKey", PGP_Name)
    
    'Need to create a name for Action.xml
    'Pad the nodeStep with leading zero if neccessary
    ActionXMLPath = SaveLocation & "50_" & GUIDString & "_" & Format(nodeStep, "#00") & ".xml"
    
    XML_Doc.Save (ActionXMLPath)
    ActionListNames = ActionListNames & ActionXMLPath & ";"
    
    nodeStep = nodeStep + 1
    '----------------------------------------------------------------------------------------------------------------------------------------
End Sub
