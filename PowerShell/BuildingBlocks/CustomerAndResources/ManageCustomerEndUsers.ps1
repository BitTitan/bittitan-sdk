# This script illustrates how to manage customer end users

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve a customer and get a ticket for it
$customerId = [GUID](Read-Host -Prompt 'Customer ID')    
$customer = Get-BT_Customer -Ticket $ticket -Id $customerId
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Create a new end user under that customer
$newEndUser = Add-BT_CustomerEndUser -Ticket $customerTicket -PrimaryEmailAddress "user@test.com" -FirstName John -LastName Smith
Write-Output $newEndUser

# Retrieve the end users under that customer
$endUsers = Get-BT_CustomerEndUser -Ticket $customerTicket
Write-Output $endUsers