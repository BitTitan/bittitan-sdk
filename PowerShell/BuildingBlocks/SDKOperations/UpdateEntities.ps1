# This script illustrates how to update entities

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Get a customer ticket
$customer = Get-BT_Customer -Ticket $ticket -CompanyName "Default"
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Retrieve and update a single endpoint
# In this example, we update the endpoint's name
$endpoint = Get-BT_Endpoint -Ticket $customerTicket -Name "Endpoint1" -IsDeleted False
$endpoint = Set-BT_Endpoint -Ticket $customerTicket -Endpoint $endpoint -Name "Endpoint1 updated"
Write-Output $endpoint

# Alternatively, the endpoint's properties can also be modified directly on the object
$endpoint.Name = "Endpoint1 updated again"
$endpoint = Set-BT_Endpoint -Ticket $customerTicket -Endpoint $endpoint
Write-Output $endpoint

# Retrieve and update multiple endpoints
$endpoints = Get-BT_Endpoint -Ticket $customerTicket -Name @("Endpoint2", "Endpoint3") -IsDeleted False
$endpoints | ForEach {
    $_.Name += " updated"
}
$endpoints = Set-BT_Endpoint -Ticket $customerTicket -Endpoint $endpoints
Write-Output $endpoints