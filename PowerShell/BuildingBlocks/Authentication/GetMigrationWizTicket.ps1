# This script shows how to get an authentication ticket for MigrationWiz

# Get a ticket for MigrationWiz
$creds = Get-Credential
$mwTicket = Get-MW_Ticket -Credentials $creds