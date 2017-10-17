# Initialize the session
.\Init.ps1

# Retrieve a customer and create an extended property associated to the customer
$customer = Get-BT_Customer -Ticket $mspc.Ticket -CompanyName "Default"
$extendedProperty = Add-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -Name "ExampleProperty" -Value "Custom Value" -ReferenceEntityType "Customer" -ReferenceEntityId $customer.Id

# Retrieve the updated customer entity
$customer = Get-BT_Customer -Ticket $mspc.Ticket -CompanyName "Default"
Write-Output "The value of the custom property with name 'DefaultCustomProperty' is '$($customer.ExtendedProperties.ExampleProperty)'."

# Delete the created extended property
Remove-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -Id $extendedProperty.Id -Force