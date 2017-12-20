# This script illustrates how to delete entities

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Get a customer ticket
$customer = Get-BT_Customer -Ticket $ticket -CompanyName "Default"
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Retrieve and delete a single endpoint
$endpoint = Get-BT_Endpoint -Ticket $customerTicket -Name "Test Endpoint" -IsDeleted False
Remove-BT_Endpoint -Ticket $customerTicket -Id $endpoint.Id

# Retrieve and delete multiple endpoints
$endpoints = Get-BT_Endpoint -Ticket $customerTicket -Name @("Endpoint 1", "Endpoint 2") -IsDeleted False
Remove-BT_Endpoint -Ticket $customerTicket -Id $endpoints.Id