# This script shows how to deploy a runbook to a customer

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Initialize variables
$offeringMetadataId = [GUID](Read-Host -Prompt 'Runbook ID')
$customerId = [GUID](Read-Host -Prompt 'Customer ID')    

# Retrieve the customer to deploy the runbook to
# And get a ticket for it
$customer = Get-BT_Customer -Ticket $ticket -Id $customerId
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Deploy the runbook
Add-BT_OfferingInstance -Ticket $customerTicket -OfferingMetadataId $offeringMetadataId
Write-Verbose "Deployed runbook: $offeringMetadataId to customer: $($customer.CompanyName)."