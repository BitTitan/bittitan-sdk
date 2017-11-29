# This script shows how to add an item to a MigrationWiz project

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$mwTicket = Get-MW_Ticket -Credentials $creds

# Retrieve an existing project
$connector = Get-MW_MailboxConnector -Ticket $mwTicket -Name "TestProject"

# Create a migration item
# Username and password are not required if the project (connector) is using admin credentials for the migration
$mailbox = Add-MW_Mailbox -Ticket $mwTicket -ConnectorId $connector.Id `
	-ImportEmailAddress import@email.com -ExportEmail export@email.com -ExportPassword your_export_password -ImportPassword your_import_password `
	-ExportUserName exprot@email.com -ImportUserName import@email.com
Write-Output "The item '$($mailbox.ExportEmailAddress)' was successfully created under '$($connector.Name)'."