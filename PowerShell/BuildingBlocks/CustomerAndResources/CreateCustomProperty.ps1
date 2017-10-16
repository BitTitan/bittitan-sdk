param(
  [Parameter(Mandatory=$true)]
  [string]$Name,
  [Parameter(Mandatory=$true)]
  [string]$Value,
  [Parameter(Mandatory=$true)]
  [string]$ReferenceEntityType,
  [Parameter(Mandatory=$true)]
  [string]$ReferenceEntityId)

# Initialize the session
.\Init.ps1

# Check if a custom property with the same fields already exists
# Name: The name of the custom property.
# Value: The value of the custom property.
# ReferenceEntityType: The type of the entity associated with the custom property. Examples include 'Customer', 'CustomerDevice', 'CustomerEndUser'.
# ReferenceEntityId: The id of the entity associated with the custom property.
$existingCustomProperty = Get-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -Name $name -ReferenceEntityType $ReferenceEntityType -ReferenceEntityId $ReferenceEntityId
if ( $existingCustomProperty -ne $null ) {
    # Write a warning if a custom property already exists
    Write-warning "Custom property with name '$name' already exists for $ReferenceEntityType entity with id: $ReferenceEntityId. A new custom property was not created."
}
else {
    # Add a new custom property
    # Use Get-Help Add-BT_ExtendedProperty to see all parameters
    Add-BT_ExtendedProperty -Ticket $mspc.WorkgroupTicket -Name $name -Value $Value -ReferenceEntityType $ReferenceEntityType -ReferenceEntityId $ReferenceEntityId
}