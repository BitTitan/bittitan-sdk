# This script shows how to get an authentication ticket
# 1) Scoped to a customer
# 2) Scoped to a workgroup

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $credentials -ServiceType BitTitan

# 1) Get a ticket for a customer
$customerId = [GUID](Read-Host -Prompt 'Customer ID')  
$customer = Get-BT_Customer -Ticket $ticket -Id $customerId
$customerTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customer.OrganizationId

# 2) Get a ticket for a workgroup
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')  
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $workgroupId
$workgroupTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $workgroup.WorkgroupOrganizationId