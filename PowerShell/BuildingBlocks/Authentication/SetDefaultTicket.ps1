# This script shows how to set a ticket to be used by default

# Get a ticket and set it as default
$creds = Get-Credential
Get-BT_Ticket -Credentials $creds -ServiceType BitTitan -SetDefault

# Retrieve customers without specifying the ticket
# This will use the ticket previously set as default
$customers = Get-BT_Customer