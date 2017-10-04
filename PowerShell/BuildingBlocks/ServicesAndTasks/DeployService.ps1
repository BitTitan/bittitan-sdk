param(
  [Parameter(Mandatory=$true)]
  [string]$ServiceKeyName)

# Initialize the session
.\Init.ps1 -Customer

# Find the service with the specified name
$offeringMetadata = Get-BT_OfferingMetadata -Ticket $mspc.Ticket -KeyName $ServiceKeyName

# Deploy the service
Add-BT_OfferingInstance -Ticket $mspc.CustomerTicket -OfferingMetadataId $offeringMetadata.Id
Write-Verbose "Deployed service: $ServiceKeyName."