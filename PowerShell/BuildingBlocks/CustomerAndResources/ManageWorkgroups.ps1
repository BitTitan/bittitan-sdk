# This script illustrates how to retrieve and create workgroups

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve workgroups
$workgroups = Get-BT_Workgroup -Ticket $ticket
Write-Output $workgroups

# Create a new workgroup
$newWorkgroup = Add-BT_Workgroup -Ticket $ticket -Name "Test Workgroup"
Write-Output $newWorkgroup