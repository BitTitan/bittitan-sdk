# This script illustrates how to retrieve feeds for a given workgroup

# Authenticate
$creds = Get-Credential

# Initialize variables
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')    

# Retrieve the top 10 most recent feeds under the workgroup
$ticketWithWorkgroupId = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan -WorkgroupId $workgroupId
$feedInstances = Get-BT_FeedInstance -Ticket $ticketWithWorkgroupId -PageSize 10 -SortBy_Created_Descending

# Retrieve the corresponding feed metadata
$feedMetadata = Get-BT_FeedMetadata -Ticket $ticketWithWorkgroupId -Id $feedInstances.FeedMetadataId

# Display the feeds
ForEach ($feedInstance in $feedInstances) {
    # Retrieve the metadata of the current feed instance
    $metadata = $feedMetadata | Where {
        $feedInstance.FeedMetadataId -eq $_.Id
    }
    
    # Display the feed keyname and parameters
    $parameterStrings = @()
    For ($i=0; $i -lt $feedInstance.Parameters.Length; $i++) {
       $parameterStrings +=  "$($feedInstance.Parameters[$i].Name)=$($feedInstance.Parameters[$i].Value)"
    }
    Write-Output "$($feedInstance.EventDate) - $($metadata.KeyName): $([string]::Join(", ", $parameterStrings))"
}
