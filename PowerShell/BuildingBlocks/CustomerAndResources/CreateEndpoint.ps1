# This script illustrates how to create an endpoint

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve a customer and get a ticket for it
$customerId = [GUID](Read-Host -Prompt 'Customer ID')    
$customer = Get-BT_Customer -Ticket $ticket -Id $customerId
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Initialize a configuration (in this case and Exchange configuration)
$endpointConfiguraton = New-BT_ExchangeConfiguration -AdministrativeUsername "TestUserName" -AdministrativePassword "TestPassword" -UseAdministrativeCredentials $True -ExchangeItemType Mail

# Create a new endpoint
$endpoint = Add-BT_Endpoint -Ticket $customerTicket -Configuration $endpointConfiguraton -Type ExchangeOnline2 -Name "Test Endpoint"
Write-Output $endpoint