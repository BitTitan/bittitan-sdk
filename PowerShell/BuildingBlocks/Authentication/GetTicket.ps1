# This script shows how to get an authentication ticket

# Get a ticket
$creds = Get-Credential
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan