param(
    [Parameter(Mandatory=$true)]
    [guid]$OfferingMetadataId,

    [Parameter(Mandatory=$true)]
    [string]$FolderPath)

# Initialize the context
.\Init.ps1

# Export the Runbook to FolderPath
Export-BT_OfferingMetadata -Ticket $mspc.Ticket -OfferingMetadataId $OfferingMetadataId -FolderPath $FolderPath