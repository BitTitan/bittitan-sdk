param(
  [Parameter(Mandatory=$true)]
  [string]$PrimaryDomain,
  [Parameter(Mandatory=$true)]
  [string]$CompanyName,
  [Parameter(Mandatory=$false)]
  [string]$CountryName = "")

# Initialize the session
.\Init.ps1

# Check if a customer with the same domain already exists
$existingCustomer = Get-BT_Customer -Ticket $mspc.Ticket -WorkgroupId $mspc.Workgroup.Id -PrimaryDomain $PrimaryDomain
if ( $existingCustomer -ne $null ) {
    # Write a warning if a customer already exists
    Write-warning "Customer with primary domain $primaryDomain already exists. A new customer was not created."
}
else {
    # Add a new customer
    # Additional parameters can be set such as CityName, IndustryType, CompanySize and more
    # Use Get-Help Add-BT_Customer to see all parameters
    Add-BT_Customer -Ticket $mspc.Ticket -WorkgroupId $mspc.Workgroup.Id -PrimaryDomain $PrimaryDomain -CompanyName $CompanyName -CountryName $CountryName
}