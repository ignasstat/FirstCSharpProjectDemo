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

    # Function to add XML nodes
    function Add-XMLNode {
        param(
            [System.Xml.XmlDocument]$XML_Doc,
            [System.Xml.XmlElement]$XML_Parent,
            [string]$NodeName,
            [string]$NodeValue
        )
        $Node = $XML_Doc.CreateElement($NodeName)
        $Node.InnerText = $NodeValue
        $XML_Parent.AppendChild($Node) | Out-Null
    }

    # Create XML document and root element
    $xmlDoc = New-Object System.Xml.XmlDocument
    $root = $xmlDoc.CreateElement("EFT_Action")
    $xmlDoc.AppendChild($root) | Out-Null

    # Common XML node creation process
    function PopulateXML {
        Add-XMLNode $xmlDoc $root "ProcessPriority" "50"
        Add-XMLNode $xmlDoc $root "GUID" $GUIDString
        Add-XMLNode $xmlDoc $root "StepSequence" ([string]$NodeStep.Value)
        Add-XMLNode $xmlDoc $root "CreationDTS" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Add-XMLNode $xmlDoc $root "RetriesRemaining" "2"
        Add-XMLNode $xmlDoc $root "LastAttemptDTS" ""
        Add-XMLNode $xmlDoc $root "LastResult" ""
        Add-XMLNode $xmlDoc $root "FailureNotificationEmail" ""
        
        $emailNode = $root.LastChild
        Add-XMLNode $xmlDoc $emailNode "To" "DataOperationsEFT-CD@transunion.co.uk"
        Add-XMLNode $xmlDoc $emailNode "CC" "DataBureau@transunion.co.uk"
        Add-XMLNode $xmlDoc $emailNode "BCC" ""
    }

    # Folder creation
    PopulateXML
    Add-XMLNode $xmlDoc $root.LastChild "Subject" "DTP - $JobClient - $JobNumber - Folder Creation - Failed"
    Add-XMLNode $xmlDoc $root.LastChild "Message" "The DTP has failed at creating folder while decrypting $PGPFolder."
    Add-XMLNode $xmlDoc $root "ActionType" "CreateFolder"
    Add-XMLNode $xmlDoc $root "DestinationFolder" $PGPFolder
    SaveXML 'Folder Creation'

    # Copy to temporary decryption folder
    PopulateXML
    Add-XMLNode $xmlDoc $root.LastChild "Subject" "DTP - $JobClient - $JobNumber - File Transfer - Failed"
    Add-XMLNode $xmlDoc $root.LastChild "Message" "The DTP has failed before decryption to copy $FileName from $SourceFolder to $PGPFolder."
    Add-XMLNode $xmlDoc $root "ActionType" "Copy"
    Add-XMLNode $xmlDoc $root "SourceFolder" $SourceFolder
    Add-XMLNode $xmlDoc $root "SourceFile" $FileName
    Add-XMLNode $xmlDoc $root "DestinationFolder" $PGPFolder
    SaveXML 'File Copy'

    # Decryption Action
    PopulateXML
    Add-XMLNode $xmlDoc $root.LastChild "Subject" "DTP - $JobClient - $JobNumber - File Decryption - Failed"
    Add-XMLNode $xmlDoc $root.LastChild "Message" "The DTP has failed at decrypting $FileName in $PGPFolder."
    Add-XMLNode $xmlDoc $root "ActionType" "Decrypt"
    Add-XMLNode $xmlDoc $root "SourceFolder" $PGPFolder
    Add-XMLNode $xmlDoc $root "SourceFile" $FileName
    Add-XMLNode $xmlDoc $root "EncryptionKey" $PGP_Name
    SaveXML 'Decryption'

    function SaveXML([string]$ActionType) {
        $ActionXMLPath = Join-Path -Path $SaveLocation -ChildPath ("50_$GUIDString_" + ('{0:D2}' -f $NodeStep.Value) + ".xml")
        $xmlDoc.Save($ActionXMLPath)
        $ActionListNames.Value += "$ActionXMLPath;"
        $NodeStep.Value++
        $xmlDoc.RemoveAll()
        $root = $xmlDoc.CreateElement("EFT_Action")
        $xmlDoc.AppendChild($root) | Out-Null
    }
}


function Create-CopyNode {
    param(
        [string]$GUIDString,
        [ref]$ActionListNames,
        [string]$SaveLocation,
        [ref]$NodeStep,
        [string]$SourceFolder,
        [string]$DestinationFolder,
        [string]$FileToCopy,
        [string]$JobNumber,
        [string]$JobClient
    )

    # Function to add XML nodes
    function Add-XMLNode {
        param(
            [System.Xml.XmlDocument]$XML_Doc,
            [System.Xml.XmlElement]$XML_Parent,
            [string]$NodeName,
            [string]$NodeValue
        )
        $Node = $XML_Doc.CreateElement($NodeName)
        $Node.InnerText = $NodeValue
        $XML_Parent.AppendChild($Node) | Out-Null
    }

    # Create XML document and root element
    $xmlDoc = New-Object System.Xml.XmlDocument
    $root = $xmlDoc.CreateElement("EFT_Action")
    $xmlDoc.AppendChild($root) | Out-Null

    # Adding common node elements
    Add-XMLNode $xmlDoc $root "ProcessPriority" "50"
    Add-XMLNode $xmlDoc $root "GUID" $GUIDString
    Add-XMLNode $xmlDoc $root "StepSequence" ([string]$NodeStep.Value)
    Add-XMLNode $xmlDoc $root "CreationDTS" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Add-XMLNode $xmlDoc $root "RetriesRemaining" "2"
    Add-XMLNode $xmlDoc $root "LastAttemptDTS" ""
    Add-XMLNode $xmlDoc $root "LastResult" ""
    Add-XMLNode $xmlDoc $root "FailureNotificationEmail" ""

    # Populate FailureNotificationEmail with email details
    $emailNode = $root.LastChild
    Add-XMLNode $xmlDoc $emailNode "To" "DataOperationsEFT-CD@transunion.co.uk"
    Add-XMLNode $xmlDoc $emailNode "CC" "DataBureau@transunion.co.uk"
    Add-XMLNode $xmlDoc $emailNode "BCC" ""
    Add-XMLNode $xmlDoc $emailNode "Subject" "DTP - $JobClient - $JobNumber - File Transfer - Failed"
    Add-XMLNode $xmlDoc $emailNode "Message" "The DTP has failed to copy $FileToCopy`r`nFrom: $SourceFolder`r`nTo: $DestinationFolder"

    # Add copy action details
    Add-XMLNode $xmlDoc $root "ActionType" "Copy"
    Add-XMLNode $xmlDoc $root "SourceFolder" $SourceFolder
    Add-XMLNode $xmlDoc $root "SourceFile" $FileToCopy
    Add-XMLNode $xmlDoc $root "DestinationFolder" $DestinationFolder
    Add-XMLNode $xmlDoc $root "DestinationFile" ""

    # Save the XML document
    $formattedNodeStep = "{0:D2}" -f $NodeStep.Value
    $ActionXMLPath = Join-Path -Path $SaveLocation -ChildPath ("50_$GUIDString_$formattedNodeStep.xml")
    $xmlDoc.Save($ActionXMLPath)
    $ActionListNames.Value += "$ActionXMLPath;"

    # Increment the step
    $NodeStep.Value++
}

# Example usage
$ActionListNames = ""
$NodeStep = 0
Create-CopyNode -GUIDString "12345" -ActionListNames ([ref]$ActionListNames) -SaveLocation "C:\Path\To\Save" -NodeStep ([ref]$NodeStep) -SourceFolder "C:\Source" -DestinationFolder "C:\Destination" -FileToCopy "example.txt" -JobNumber "001" -JobClient "ClientX"

function Create-DeletionNode {
    param(
        [string]$GUIDString,
        [ref]$ActionListNames,
        [string]$SaveLocation,
        [ref]$NodeStep,
        [string]$FolderToDelete,
        [string]$JobNumber,
        [string]$JobClient
    )

    # Function to add XML nodes
    function Add-XMLNode {
        param(
            [System.Xml.XmlDocument]$XML_Doc,
            [System.Xml.XmlElement]$XML_Parent,
            [string]$NodeName,
            [string]$NodeValue
        )
        $Node = $XML_Doc.CreateElement($NodeName)
        $Node.InnerText = $NodeValue
        $XML_Parent.AppendChild($Node) | Out-Null
    }

    # Create XML document and root element
    $xmlDoc = New-Object System.Xml.XmlDocument
    $root = $xmlDoc.CreateElement("EFT_Action")
    $xmlDoc.AppendChild($root) | Out-Null

    # Common properties for the action
    Add-XMLNode $xmlDoc $root "ProcessPriority" "50"
    Add-XMLNode $xmlDoc $root "GUID" $GUIDString
    Add-XMLNode $xmlDoc $root "StepSequence" ([string]$NodeStep.Value)
    Add-XMLNode $xmlDoc $root "CreationDTS" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Add-XMLNode $xmlDoc $root "RetriesRemaining" "2"
    Add-XMLNode $xmlDoc $root "LastAttemptDTS" ""
    Add-XMLNode $xmlDoc $root "LastResult" ""
    Add-XMLNode $xmlDoc $root "FailureNotificationEmail" ""

    # Failure Notification Email details
    $emailNode = $root.LastChild
    Add-XMLNode $xmlDoc $emailNode "To" "DataOperationsEFT-CD@transunion.co.uk"
    Add-XMLNode $xmlDoc $emailNode "CC" "DataBureau@transunion.co.uk"
    Add-XMLNode $xmlDoc $emailNode "BCC" ""
    Add-XMLNode $xmlDoc $emailNode "Subject" "DTP failure - $JobClient - $JobNumber"
    Add-XMLNode $xmlDoc $emailNode "Message" "The DTP has failed to delete folder $FolderToDelete."

    # Deletion specific properties
    Add-XMLNode $xmlDoc $root "ActionType" "DeleteFolder"
    Add-XMLNode $xmlDoc $root "SourceFolder" $FolderToDelete
    Add-XMLNode $xmlDoc $root "SourceFile" ""
    Add-XMLNode $xmlDoc $root "DestinationFolder" ""
    Add-XMLNode $xmlDoc $root "DestinationFile" ""

    # Save the XML to a file
    $formattedNodeStep = "{0:D2}" -f $NodeStep.Value
    $ActionXMLPath = Join-Path -Path $SaveLocation -ChildPath ("50_$GUIDString_$formattedNodeStep.xml")
    $xmlDoc.Save($ActionXMLPath)
    $ActionListNames.Value += "$ActionXMLPath;"

    # Increment the step
    $NodeStep.Value++
}

# Example usage:
$ActionListNames = ""
$NodeStep = 0
Create-DeletionNode -GUIDString "12345" -ActionListNames ([ref]$ActionListNames) -SaveLocation "C:\Path\To\Save" -NodeStep ([ref]$NodeStep) -FolderToDelete "C:\Path\To\Delete" -JobNumber "001" -JobClient "ClientX"

function Create-MoveNode {
    param(
        [string]$GUIDString,
        [ref]$ActionListNames,
        [string]$SaveLocation,
        [ref]$NodeStep,
        [string]$SourceFolder,
        [string]$DestinationFolder,
        [string]$FileToCopy,
        [string]$JobNumber,
        [string]$JobClient
    )

    # Function to add XML nodes
    function Add-XMLNode {
        param(
            [System.Xml.XmlDocument]$XML_Doc,
            [System.Xml.XmlElement]$XML_Parent,
            [string]$NodeName,
            [string]$NodeValue
        )
        $Node = $XML_Doc.CreateElement($NodeName)
        $Node.InnerText = $NodeValue
        $XML_Parent.AppendChild($Node) | Out-Null
    }

    # Create XML document and root element
    $xmlDoc = New-Object System.Xml.XmlDocument
    $root = $xmlDoc.CreateElement("EFT_Action")
    $xmlDoc.AppendChild($root) | Out-Null

    # Common properties for the action
    Add-XMLNode $xmlDoc $root "ProcessPriority" "50"
    Add-XMLNode $xmlDoc $root "GUID" $GUIDString
    Add-XMLNode $xmlDoc $root "StepSequence" ([string]$NodeStep.Value)
    Add-XMLNode $xmlDoc $root "CreationDTS" (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    Add-XMLNode $xmlDoc $root "RetriesRemaining" "2"
    Add-XMLNode $xmlDoc $root "LastAttemptDTS" ""
    Add-XMLNode $xmlDoc $root "LastResult" ""
    Add-XMLNode $xmlDoc $root "FailureNotificationEmail" ""

    # Failure Notification Email details
    $emailNode = $root.LastChild
    Add-XMLNode $xmlDoc $emailNode "To" "DataOperationsEFT-CD@transunion.co.uk"
    Add-XMLNode $xmlDoc $emailNode "CC" "DataBureau@transunion.co.uk"
    Add-XMLNode $xmlDoc $emailNode "BCC" ""
    Add-XMLNode $xmlDoc $emailNode "Subject" "DTP - $JobClient - $JobNumber - File Transfer - Failed"
    Add-XMLNode $xmlDoc $emailNode "Message" "The DTP has failed to move $FileToCopy`r`nFrom: $SourceFolder`r`nTo: $DestinationFolder"

    # Move action specific properties
    Add-XMLNode $xmlDoc $root "ActionType" "Move"
    Add-XMLNode $xmlDoc $root "SourceFolder" $SourceFolder
    Add-XMLNode $xmlDoc $root "SourceFile" $FileToCopy
    Add-XMLNode $xmlDoc $root "DestinationFolder" $DestinationFolder
    Add-XMLNode $xmlDoc $root "DestinationFile" ""

    # Save the XML document
    $formattedNodeStep = "{0:D2}" -f $NodeStep.Value
    $ActionXMLPath = Join-Path -Path $SaveLocation -ChildPath ("50_$GUIDString_$formattedNodeStep.xml")
    $xmlDoc.Save($ActionXMLPath)
    $ActionListNames.Value += "$ActionXMLPath;"

    # Increment the step
    $NodeStep.Value++
}

# Example usage:
$ActionListNames = ""
$NodeStep = 0
Create-MoveNode -GUIDString "12345" -ActionListNames ([ref]$ActionListNames) -SaveLocation "C:\Path\To\Save" -NodeStep ([ref]$NodeStep) -SourceFolder "C:\Path\To\Source" -DestinationFolder "C:\Path\To\Destination" -FileToCopy "example.txt" -JobNumber "001" -JobClient "ClientX"
