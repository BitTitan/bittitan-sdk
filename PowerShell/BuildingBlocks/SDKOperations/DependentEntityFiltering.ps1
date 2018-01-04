# This script illustrates how to retrieve entities with dependent entity filtering

# To get a workgroup that has a customer named BitTitan
$workgroup = Get-BT_Workgroup -Ticket $ticket -CustomerCompanyName 'BitTitan'

# To get all customer end users with unprocessed subscriptions
$customerEndUsers = Get-BT_CustomerEndUser -Ticket $ticket -SubscriptionSubscriptionProcessState NotProcessed