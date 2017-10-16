# Initialize the session
.\Init.ps1

# Retrieve a customer and create an extended property associated to the customer
$customer = Get-BT_Customer -Ticket $mspc.Ticket -CompanyName "Default"
$extendedProperty = Add-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -Name "ExampleProperty" -Value "Custom Value" -ReferenceEntityType "Customer" -ReferenceEntityId $customer.Id
Write-Output "Old value is: $($extendedProperty.Value)."

# Update the custom property
$extendedProperty.Value = "New Custom Value"
$extendedProperty = Set-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -extendedproperty $extendedProperty
Write-Output "New value is: $($extendedProperty.Value)."

# Delete the created custom property
Remove-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -Id $extendedProperty.Id -Force