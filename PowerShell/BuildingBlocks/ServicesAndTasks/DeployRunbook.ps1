# Initialize variables
$offeringMetadataId = Read-Host -Prompt 'Runbook ID'
$customerId = [GUID](Read-Host -Prompt 'Customer ID')    

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve the customer to deploy the runbook to
# And get a ticket for it
$customer = Get-BT_Customer -Ticket $ticket -Id $customerId
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Deploy the runbook
Add-BT_OfferingInstance -Ticket $customerTicket -OfferingMetadataId $offeringMetadataId
Write-Warning "Deployed runbook: $offeringMetadataId to customer: $($customer.CompanyName)."