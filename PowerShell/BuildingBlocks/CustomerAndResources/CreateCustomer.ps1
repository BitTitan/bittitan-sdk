# This script illustrates how to create a customer

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Initialize variables
$primaryDomain = "an-it-company.com"
$companyName = "IT Company"
$countryName = "Italy" # Optional

# Retrieve the workgroup that the customer should be created under
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')    
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $workgroupId

# Check if a customer with the same domain already exists under this workgroup
$existingCustomer = Get-BT_Customer -Ticket $ticket -WorkgroupId $workgroup.Id -PrimaryDomain $primaryDomain
if ( $existingCustomer -ne $null ) {
    # Write a warning if a customer already exists
    Write-warning "Customer with primary domain $primaryDomain already exists. A new customer was not created."
}
else {
    # Add a new customer
    # Additional parameters can be set such as CityName, IndustryType, CompanySize and more
    # Use Get-Help Add-BT_Customer to see all parameters
    Add-BT_Customer -Ticket $ticket -WorkgroupId $workgroup.Id -PrimaryDomain $primaryDomain -CompanyName $companyName -CountryName $countryName
}