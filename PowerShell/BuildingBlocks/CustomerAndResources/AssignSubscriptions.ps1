# Initialize the context
.\Init.ps1 -Customer

# Get the end users under that customer (this example retrieves the top 100)
$customerEndUsers =  Get-BT_CustomerEndUser -Ticket $mspc.CustomerTicket -OrganizationId $mspc.Customer.OrganizationId -IsDeleted $False -PageSize 100
if ( $customerEndUsers ) {
    # Check for existing subscriptions
    $existingSubscriptions = Get-BT_Subscription -Ticket $mspc.WorkgroupTicket -SubscriptionEntityReferenceType "CustomerEndUser" -EntityReferenceId $customerEndUsers.Id -IsDeleted $False
    
    # Filter out end users who already have a subscription
    $customerEndUsersToSubscribe = $customerEndUsers | Where {
        !( $existingSubscriptions.EntityReferenceId -contains $_.Id ) 
    }
    
    # Get the product sku id for the MSPC yearly subscription
    $productSkuId = Get-BT_ProductSkuId -Ticket $mspc.Ticket -ProductName MspcEndUserYearlySubscription

    # Assign subscriptions to each customer end user
    $customerEndUsersToSubscribe | ForEach {
        Add-BT_Subscription -Ticket $mspc.WorkgroupTicket -SubscriptionEntityReferenceType CustomerEndUser -EntityReferenceId $_.Id -ProductSkuId $productSkuId 
    }
}