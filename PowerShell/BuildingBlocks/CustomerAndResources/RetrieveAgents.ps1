# This script illustrates how to retrieve agents in a workgroup

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve a workgroup and get a ticket for it
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')    
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $workgroupId
$workgroupTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $workgroup.WorkgroupOrganizationId

# Retrieve the IDs of the agents in the workgroup
$agentIDs = (Get-BT_OrganizationMember -Ticket $workgroupTicket).SystemUserId
Write-Output $agentIDs