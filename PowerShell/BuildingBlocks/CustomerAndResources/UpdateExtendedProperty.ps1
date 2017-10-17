# Initialize the session
.\Init.ps1

# Retrieve a customer and create an extended property associated to the customer
$customer = Get-BT_Customer -Ticket $mspc.Ticket -CompanyName "Default"
$extendedProperty = Add-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -Name "ExampleProperty" -Value "Custom Value" -ReferenceEntityType "Customer" -ReferenceEntityId $customer.Id
Write-Output "Value of Extended Property is: $($extendedProperty.Value)."

# Update the extended property
$extendedProperty.Value = "New Custom Value"
$extendedProperty = Set-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -ExtendedProperty $extendedProperty
Write-Output "Updated value of Extended Property is: $($extendedProperty.Value)."

# Delete the created extended property
Remove-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -Id $extendedProperty.Id -Force