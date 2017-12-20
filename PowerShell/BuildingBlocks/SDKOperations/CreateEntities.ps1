# This script illustrates how to create entities

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Create a customer
$customer = Add-BT_Customer -Ticket $ticket -CompanyName "TestCustomer" -PrimaryDomain "testCustomer.com"
Write-Output $customer