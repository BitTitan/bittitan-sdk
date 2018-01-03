# This script shows how to import users from a CSV and assign subscriptions to them

# Get credentials
$creds = Get-Credential -Message "Enter BitTitan credentials"

# Get ticket
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan -Environment "BT"

# Get customer and its workgroup
$customerId = [GUID](Read-Host -Prompt 'Customer ID')   
$customer = Get-BT_Customer -Ticket $ticket -Id $customerId
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $customer.WorkgroupId

# Get customer and workgroup tickets
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId
$workgroupTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $workgroup.WorkgroupOrganizationId     

# Get CSV path
$CSVPath = Read-Host "Please enter the path of your CSV file (example: C:\scripts\test.csv)"

# Get the product sku id for the MSPC yearly subscription
$productSkuId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription

# Import the CSV and process each line
# Variable names correspond to the column names in the csv
Import-Csv -Path $CSVPath | ForEach {
   
    # Create customer end user
    $customerEndUser = Add-BT_CustomerEndUser -Ticket $customerTicket -PrimaryEmailAddress $_.PrimaryEmailAddress -FirstName $_.FirstName -LastName $_.LastName -PrimaryIdentity $_.PrimaryIdentity
    
    # Assign subscription
    Add-BT_Subscription -Ticket $workgroupTicket -SubscriptionEntityReferenceType CustomerEndUser -EntityReferenceId $customerEndUser.Id -ProductSkuId $productSkuId 
} 
