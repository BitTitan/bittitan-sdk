# This script illustrates how to filter entities when retrieving them

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan
$mwTicket = Get-MW_Ticket -Credentials $creds 

# Retrieve endpoints, filtered by date and string property
$endpoints = Get-BT_Endpoint -Ticket $ticket -Created @("> 01-01-2016", "< 01-01-2018") -Name @("Endpoint1", "Endpoint2") 

# Retrieve mailbox connectors, filtered by a string with a wildcard and by multiple Ids
$mailboxConnectors = Get-MW_MailboxConnector -Ticket $mwTicket -FilterBy_String_Name "%project" -SelectedExportEndpointId $endpoints.Id