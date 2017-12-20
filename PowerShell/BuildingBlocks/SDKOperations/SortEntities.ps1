# This script illustrates how to sort entities when retrieving them

# Authenticate
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan

# Retrieve endpoints, sorted by created date
$endpoints = Get-BT_Endpoint -Ticket $ticket -SortBy_Created_Descending