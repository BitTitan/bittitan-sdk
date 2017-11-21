# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan
     
# Retrieve a customer and its workgroup
$customerId = [GUID](Read-Host -Prompt 'Customer ID')    
$customer = Get-BT_Customer -Ticket $ticket -Id $customerId
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $customer.WorkgroupId

# Initialize additional tickets
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId
$workgroupTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $workgroup.WorkgroupOrganizationId        

# Get the end users under that customer (this example retrieves the top 100)
$customerEndUsers =  Get-BT_CustomerEndUser -Ticket $customerTicket -IsDeleted $False -PageSize 100
if ( $customerEndUsers ) {
    # Check for existing subscriptions
    $existingSubscriptions = Get-BT_Subscription -Ticket $workgroupTicket -SubscriptionEntityReferenceType CustomerEndUser -EntityReferenceId $customerEndUsers.Id -IsDeleted $False
    
    # Filter out end users who already have a subscription
    $customerEndUsersToSubscribe = $customerEndUsers | Where {
        $existingSubscriptions.EntityReferenceId -notcontains $_.Id 
    }
    
    # Get the product sku id for the MSPC yearly subscription
    $productSkuId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription

    # Assign subscriptions to each customer end user
    $customerEndUsersToSubscribe | ForEach {
        Add-BT_Subscription -Ticket $workgroupTicket -SubscriptionEntityReferenceType CustomerEndUser -EntityReferenceId $_.Id -ProductSkuId $productSkuId 
    }
    $assignedSubscriptionCount = $customerEndUsersToSubscribe.Length
    Write-Verbose "Successfully assigned subscription to $assignedSubscriptionCount end users."
}