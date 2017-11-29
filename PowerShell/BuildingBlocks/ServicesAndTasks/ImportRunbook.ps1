# This script shows how to import a runbook to a workgroup

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Initialize variables
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID to import to')
$filePath = Read-Host -Prompt 'File path to import from'

# Get a ticket scoped to the workgroup
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $workgroupId
$workgroupTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $workgroup.WorkgroupOrganizationId

# Import the runbook from FilePath
Import-BT_OfferingMetadata -Ticket $workgroupTicket -FilePath $filePath