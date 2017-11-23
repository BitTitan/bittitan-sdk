# Initialize variables
$offeringMetadataId = Read-Host -Prompt 'Runbook ID'
$folderPath = Read-Host -Prompt 'Folder path to export to'

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Export the runbook to FolderPath
Export-BT_OfferingMetadata -Ticket $ticket -OfferingMetadataId $offeringMetadataId -FolderPath $folderPath