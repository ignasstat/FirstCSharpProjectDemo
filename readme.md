function Add-XMLNode {
    param (
        [System.Xml.XmlDocument]$XML_Doc,
        [System.Xml.XmlElement]$XML_Parent,
        [string]$NodeName,
        [string]$NodeValue
    )
    $Node = $XML_Doc.CreateElement($NodeName)
    $Node.InnerText = $NodeValue
    $XML_Parent.AppendChild($Node) | Out-Null
}

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

    $XML_Child = $XML_Action.LastChild
    Add-XMLNode $xmlDoc $XML_Child "To" "DataOperationsEFT-CD@transunion.co.uk"
    Add-XMLNode $xmlDoc $XML_Child "CC" "DataBureau@transunion.co.uk"
    Add-XMLNode $xmlDoc $XML_Child "BCC" ""
    Add-XMLNode $xmlDoc $XML_Child "Subject" "DTP - $JobClient - $JobNumber - Folder Creation - Failed"
    Add-XMLNode $xmlDoc $XML_Child "Message" "The DTP has failed at creating folder while decrypting $PGPFolder."

    Save-XML $xmlDoc $SaveLocation $GUIDString $NodeStep

    Reset-XMLDoc $xmlDoc $XML_Action
    Add-XMLNode $xmlDoc $XML_Action "ActionType" "Copy"
    Add-XMLNode $xmlDoc $XML_Action "SourceFolder" $SourceFolder
    Add-XMLNode $xmlDoc $XML_Action "SourceFile" $FileName
    Add-XMLNode $xmlDoc $XML_Action "DestinationFolder" $PGPFolder
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

$ActionListNames = ""
$NodeStep = 0
$saveFolder = "\\cig.local\data\AppData\SFTP\Data\Usr\DataBureau\Configuration\Scripts\Test\CallTrace Console\CTC124_125\Other"
$pgpFolder = "\\cig.local\data\AppData\SFTP\Data\Usr\DataBureau\Configuration\Scripts
