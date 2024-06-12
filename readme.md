function Add-XMLNode {
    param (
        [System.Xml.XmlDocument]$XML_Doc,     # Document where nodes are created
        [System.Xml.XmlElement]$XML_Parent,   # Parent element to append this node
        [string]$NodeName,                    # New node's name
        [string]$NodeValue                    # New node's value
    )
    
    # Create the new element in the document
    $Node = $XML_Doc.CreateElement($NodeName)
    $Node.InnerText = $NodeValue

    # Append the newly created element to the parent
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

    # Initialize XML document
    $xmlDoc = New-Object System.Xml.XmlDocument
    $root = $xmlDoc.CreateElement("EFT_Action")
    $xmlDoc.AppendChild($root) | Out-Null


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

    Add-XMLNode $xmlDoc $XML_Action "ActionType" "CreateFolder"
    Add-XMLNode $xmlDoc $XML_Action "DestinationFolder" $PGPFolder

    Save-XML $xmlDoc $SaveLocation $GUIDString $NodeStep

    # Second XML action: Copy Files
    Reset-XMLDoc $xmlDoc $XML_Action
    Add-XMLNode $xmlDoc $XML_Action "ProcessPriority" "50"
    # Additional XML nodes follow a similar pattern as above, adjusting attributes accordingly
    # Populate nodes as required
    Add-XMLNode $xmlDoc $XML_Action "ActionType" "Copy"
    Add-XMLNode $xmlDoc $XML_Action "SourceFolder" $SourceFolder
    Add-XMLNode $xmlDoc $XML_Action "SourceFile" $FileName
    Add-XMLNode $xmlDoc $XML_Action "DestinationFolder" $PGPFolder

    Save-XML $xmlDoc $SaveLocation $GUIDString $NodeStep

    # Third XML action: Decrypt Files
    Reset-XMLDoc $xmlDoc $XML_Action
    # Populate XML nodes
    Add-XMLNode $xmlDoc $XML_Action "ActionType" "Decrypt"
    Add-XMLNode $xmlDoc $XML_Action "SourceFolder" $PGPFolder
    Add-XMLNode $xmlDoc $XML_Action "SourceFile" $FileName
    Add-XMLNode $xmlDoc $XML_Action "EncryptionKey" $PGP_Name

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


$ActionListNames = ""
$NodeStep = 0
$saveFolder = "\\cig.local\data\AppData\SFTP\Data\Usr\DataBureau\Configuration\Scripts\Test\CallTrace Console\CTC124_125\Other"
$pgpF = "\\cig.local\data\AppData\SFTP\Data\Usr\DataBureau\Configuration\Scripts\Test\CallTrace Console\CTC124_125\Other\PGPFolder"
Create-DecryptionNode -GUIDString "0001" -ActionListNames ([ref]$ActionListNames) -SaveLocation $saveFolder -NodeStep ([ref]$NodeStep) -FileName 'Test.txt' -SourceFolder $saveFolder -PGPFolder $PGPFolder -JobNumber "12345" -JobClient "ClientA"
     

You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
Reset-XMLDoc : Cannot process argument transformation on parameter 'xmlDoc'. Reference type is expected in argument.
At line:59 char:18
+     Reset-XMLDoc $xmlDoc $XML_Action
+                  ~~~~~~~
    + CategoryInfo          : InvalidData: (:) [Reset-XMLDoc], ParameterBindingArgumentTransformationException
    + FullyQualifiedErrorId : ParameterArgumentTransformationError,Reset-XMLDoc
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
Reset-XMLDoc : Cannot process argument transformation on parameter 'xmlDoc'. Reference type is expected in argument.
At line:71 char:18
+     Reset-XMLDoc $xmlDoc $XML_Action
+                  ~~~~~~~
    + CategoryInfo          : InvalidData: (:) [Reset-XMLDoc], ParameterBindingArgumentTransformationException
    + FullyQualifiedErrorId : ParameterArgumentTransformationError,Reset-XMLDoc
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
 
You cannot call a method on a null-valued expression.
At line:14 char:5
+     $XML_Parent.AppendChild($Node) | Out-Null
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidOperation: (:) [], RuntimeException
    + FullyQualifiedErrorId : InvokeMethodOnNull
