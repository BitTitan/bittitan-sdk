# This script illustrates how to retrieve report data

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve a workgroup and get a ticket for it
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')    
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $workgroupId
$workgroupTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $workgroup.WorkgroupOrganizationId

# Retrieve the last generated report and its data.
$reportInstance = Get-BT_ReportInstance -Ticket $workgroupTicket -PageSize 1 -SortBy_GeneratedOn_Descending
$reportWithData = Get-BT_ReportWithData -Ticket $workgroupTicket -ReportInstanceId $reportInstance.Id
Write-Output $reportWithData