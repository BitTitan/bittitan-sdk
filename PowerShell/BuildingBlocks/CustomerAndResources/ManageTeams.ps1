# This script illustrates how to manage teams

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve a workgroup and get a ticket for it
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')    
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $workgroupId
$workgroupTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $workgroup.WorkgroupOrganizationId

# Retrieve teams in the workgroup
$teams = Get-BT_Team -Ticket $workgroupTicket
Write-Output $teams

# Create a new team under the workgroup
$team = Add-BT_Team -Ticket $workgroupTicket -Name "Test Team" -AssignmentType LeastLoad
Write-Output $team