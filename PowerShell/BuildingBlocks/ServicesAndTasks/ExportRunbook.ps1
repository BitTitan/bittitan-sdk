# This script shows how to export a runbook to a file

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Initialize variables
$offeringMetadataId = [GUID](Read-Host -Prompt 'Runbook ID')
$folderPath = Read-Host -Prompt 'Folder path to export to'

# Export the runbook to FolderPath
Export-BT_OfferingMetadata -Ticket $ticket -OfferingMetadataId $offeringMetadataId -FolderPath $folderPath