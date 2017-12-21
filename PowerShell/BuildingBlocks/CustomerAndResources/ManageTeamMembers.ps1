# This script illustrates how to manage team members

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve a workgroup and get a ticket for it
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')    
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $workgroupId
$workgroupTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $workgroup.WorkgroupOrganizationId

# Retrieve a team
$team = Get-BT_Team -Ticket $workgroupTicket -Name "Test Team"

# Retrieve the IDs of the agents in the workgroup
$agentIDs = (Get-BT_OrganizationMember -Ticket $workgroupTicket).SystemUserId

# Add one agent to the team
$newTeamMember = Add-BT_TeamMembership -Ticket $workgroupTicket -TeamId $team.Id -SystemUserId $agentIDs[0]

# Retrieve the team members
$teamMembers = Get-BT_TeamMembership -Ticket $ticket -TeamId $team.Id
Write-Output $teamMembers

# Remove the agent from the team
Remove-BT_TeamMembership -Ticket $workgroupTicket -Id $newTeamMember.Id