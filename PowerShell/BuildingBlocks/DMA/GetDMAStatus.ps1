# This script shows how to retrieve the DMA status for devices and users

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve the customer and get a ticket for it
$customerId = [GUID](Read-Host -Prompt 'Customer ID')    
$customer = Get-BT_Customer -Ticket $ticket -Id $customerId
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# Retrieve top 10 devices
$devices = Get-BT_CustomerDevice -Ticket $customerTicket -PageSize 10

# Retrieve information for each user on each device found previously
$devices | ForEach {
    # Retrieve the information for each user
    $deviceName = $_.DeviceName
    $deviceUsers = Get-BT_CustomerDeviceUser -Ticket $customerTicket -DeviceId $_.Id
    
    # Retrieve the corresponding end user information and print the agent status
    $deviceUsers | ForEach {
        $endUser = Get-BT_CustomerEndUser -Ticket $customerTicket -Id $_.EndUserId
        Write-Output "$($deviceName): $($endUser.PrimaryEmailAddress) - $($_.AgentStatus)"
    }
}
