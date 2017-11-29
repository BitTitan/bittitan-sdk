# This script illustrates how to manage extended properties on entities
# Extended properties can be added to various entities. Examples include 'Customer', 'CustomerDevice' and 'CustomerEndUser'.

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve a customer
$customer = Get-BT_Customer -Ticket $ticket -CompanyName "Default"
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Add two extended properties to this customer
$customer.ExtendedProperties.SomeProperty = 'SomeValue'
$customer.ExtendedProperties.AnotherProperty = 'AnotherValue'
$customer = Set-BT_Customer -customer $customer -Ticket $customerTicket

# Lookup the current extended properties
Write-Output "The customer has the following custom properties:"
$customer.ExtendedProperties
Write-Output "The value of the custom property with name 'SomeProperty' is '$($customer.ExtendedProperties.SomeProperty)'."

# Update one extended property, delete the other one
$customer.ExtendedProperties.SomeProperty = 'SomeNewValue'
$customer.ExtendedProperties.Remove('AnotherProperty')
$customer = Set-BT_Customer -customer $customer -Ticket $customerTicket

# Lookup the updated extended properties again
Write-Output "The customer now has the following custom properties:"
$customer.ExtendedProperties
