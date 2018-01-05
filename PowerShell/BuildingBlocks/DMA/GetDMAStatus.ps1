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

# Retrieve agent status for each device
Write-Output $devices | Select DeviceName, AgentStatus

# Retrieve information for each user on each device found previously
$deviceUserInfos = Get-Bt_CustomerDeviceUserInfo -Ticket $customerTicket

# Retrieve agent status for each user on each device
Write-Output $deviceUserInfos | Select EndUserId, DeviceId, AgentStatus