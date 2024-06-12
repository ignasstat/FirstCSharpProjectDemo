function Create-DecryptionNode {
    param(
        [string]$GUIDString,
        [ref]$ActionListNames,
        [string]$SaveLocation,
        [ref]$NodeStep,
        [string]$FileName,
        [string]$SourceFolder,
        [string]$PGPFolder,
        [string]$PGP_Name,
        [string]$JobNumber,
        [string]$JobClient
    )

    # Initialize XML document
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.async = $false
    $xmlDoc.validateOnParse = $false

    # First XML action: Create Folder
    $XML_Action = $xmlDoc.CreateElement("EFT_Action")
    $xmlDoc.AppendChild($XML_Action) | Out-Null

    XML_AddNode $xmlDoc $XML_Action "ProcessPriority" "50"
    XML_AddNode $xmlDoc $XML_Action "GUID" $GUIDString
    XML_AddNode $xmlDoc $XML_Action "StepSequence" ([string]$NodeStep.Value)
    XML_AddNode $xmlDoc $XML_Action "CreationDTS" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    XML_AddNode $xmlDoc $XML_Action "RetriesRemaining" "2"
    XML_AddNode $xmlDoc $XML_Action "LastAttemptDTS" ""
    XML_AddNode $xmlDoc $XML_Action "LastResult" ""
    XML_AddNode $xmlDoc $XML_Action "FailureNotificationEmail" ""

    $XML_Child = $XML_Action.LastChild
    XML_AddNode $xmlDoc $XML_Child "To" "DataOperationsEFT-CD@transunion.co.uk"
    XML_AddNode $xmlDoc $XML_Child "CC" "DataBureau@transunion.co.uk"
    XML_AddNode $xmlDoc $XML_Child "BCC" ""
    XML_AddNode $xmlDoc $XML_Child "Subject" "DTP - $JobClient - $JobNumber - Folder Creation - Failed"
    XML_AddNode $xmlDoc $XML_Child "Message" "The DTP has failed at creating folder while decrypting $PGPFolder."

    XML_AddNode $xmlDoc $XML_Action "ActionType" "CreateFolder"
    XML_AddNode $xmlDoc $XML_Action "DestinationFolder" $PGPFolder

    Save-XML $xmlDoc $SaveLocation $GUIDString $NodeStep

    # Second XML action: Copy Files
    Reset-XMLDoc $xmlDoc $XML_Action
    XML_AddNode $xmlDoc $XML_Action "ProcessPriority" "50"
    # Additional XML nodes follow a similar pattern as above, adjusting attributes accordingly
    # Populate nodes as required
    XML_AddNode $xmlDoc $XML_Action "ActionType" "Copy"
    XML_AddNode $xmlDoc $XML_Action "SourceFolder" $SourceFolder
    XML_AddNode $xmlDoc $XML_Action "SourceFile" $FileName
    XML_AddNode $xmlDoc $XML_Action "DestinationFolder" $PGPFolder

    Save-XML $xmlDoc $SaveLocation $GUIDString $NodeStep

    # Third XML action: Decrypt Files
    Reset-XMLDoc $xmlDoc $XML_Action
    # Populate XML nodes
    XML_AddNode $xmlDoc $XML_Action "ActionType" "Decrypt"
    XML_AddNode $xmlDoc $XML_Action "SourceFolder" $PGPFolder
    XML_AddNode $xmlDoc $XML_Action "SourceFile" $FileName
    XML_AddNode $xmlDoc $XML_Action "EncryptionKey" $PGP_Name

    Save-XML $xmlDoc $SaveLocation $GUIDString $NodeStep
}

function Reset-XMLDoc([ref]$xmlDoc, [ref]$XML_Action) {
    $xmlDoc.RemoveAll()
    $XML_Action = $xmlDoc.CreateElement("EFT_Action")
    $xmlDoc.AppendChild($XML_Action) | Out-Null
}

function Save-XML([System.Xml.XmlDocument]$xmlDoc, [string]$SaveLocation, [string]$GUIDString, [ref]$NodeStep) {
    $ActionXMLPath = Join-Path $SaveLocation ("50_" + $GUIDString + "_" + ('{0:D2}' -f $NodeStep.Value) + ".xml")
    $xmlDoc.Save($ActionXMLPath)
    $ActionListNames.Value += $ActionXMLPath + ";"
    $NodeStep.Value++
}
