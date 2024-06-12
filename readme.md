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
    $root = $xmlDoc.CreateElement("EFT_Action")
    $xmlDoc.AppendChild($root) | Out-Null

    $XML_Action = $root

    Add-XMLNode $xmlDoc $XML_Action "ProcessPriority" "50"
    Add-XMLNode $xmlDoc $XML_Action "GUID" $GUIDString
    Add-XMLNode $xmlDoc $XML_Action "StepSequence" ([string]$NodeStep.Value)
    Add-XMLNode $xmlDoc $XML_Action "CreationDTS" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Add-XMLNode $xmlDoc $XML_Action "RetriesRemaining" "2"
    Add-XMLNode $xmlDoc $XML_Action "LastAttemptDTS" ""
    Add-XMLNode $xmlDoc $XML_Action "LastResult" ""
    Add-XMLNode $xmlDoc $XML_Action "FailureNotificationEmail" ""

    $XML_Child = $xmlDoc.CreateElement("EmailDetails")
    $XML_Action.AppendChild($XML_Child) | Out-Null

    Add-XMLNode $xmlDoc $XML_Child "To" "DataOperationsEFT-CD@transunion.co.uk"
    Add-XMLNode $xmlDoc $XML_Child "CC" "DataBureau@transunion.co.uk"
    Add-XMLNode $xmlDoc $XML_Child "BCC" ""
    Add-XMLNode $xmlDoc $XML_Child "Subject" "DTP - $JobClient - $JobNumber - Folder Creation - Failed"
    Add-XMLNode $xmlDoc $XML_Child "Message" "The DTP has failed at creating folder while decrypting $PGPFolder."

    Add-XMLNode $xmlDoc $XML_Action "ActionType" "CreateFolder"
    Add-XMLNode $xmlDoc $XML_Action "DestinationFolder" $PGPFolder

    Save-XML $xmlDoc $SaveLocation $GUIDString $NodeStep

    # Subsequent actions should be similar but with appropriate modifications for the specific actions.
}

function Reset-XMLDoc([ref]$xmlDoc, [ref]$XML_Action) {
    $xmlDoc.Value.RemoveAll()
    $XML_Action.Value = $xmlDoc.Value.CreateElement("EFT_Action")
    $xmlDoc.Value.AppendChild($XML_Action.Value) | Out-Null
}

function Save-XML([System.Xml.XmlDocument]$xmlDoc, [string]$SaveLocation, [string]$GUIDString, [ref]$NodeStep) {
    $ActionXMLPath = Join-Path $SaveLocation ("50_" + $GUIDString + "_" + ('{0:D2}' -f $NodeStep.Value) + ".xml")
    $xmlDoc.Save($ActionXMLPath)
    $ActionListNames.Value += $ActionXMLPath + ";"
    $NodeStep.Value++
}
