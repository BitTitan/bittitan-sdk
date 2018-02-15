# This script illustrates how to use pipelines with the BitTitan SDK

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Get a customer ticket
$customer = Get-BT_Customer -Ticket $ticket -CompanyName "Default"
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Get all users associated to the customer and archive them
Get-BT_CustomerEndUser -Ticket $customerTicket -IsDeleted $false -RetrieveAll | Set-BT_CustomerEndUser -Ticket $customerTicket -IsArchived $true

# Get all devices associated to the customer and delete them
Get-BT_CustomerDevice -Ticket $customerTicket -IsDeleted $false -RetrieveAll | Remove-BT_CustomerDevice -Ticket $customerTicket -Force