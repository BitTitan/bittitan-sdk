# This script illustrates how to start migrations
# This assumes that the migration items have corresponding subscriptions or licenses
# Examples of how to assign subscriptions can be found in our GitHub sample scripts

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$mwTicket = Get-MW_Ticket -Credentials $creds

# Retrieve an existing project
$connector = Get-MW_MailboxConnector -Ticket $mwTicket -Name "TestProject"

# Retrieve all the migration items in the project
$mailboxes = Get-MW_Mailbox -Ticket $mwTicket -ConnectorId $connector.Id -RetrieveAll | ForEach {
    # Start a migration for each item
    # -Type indicates the type of migration, e.g. Trial, Full, etc.
    # -ItemTypes indicates what item types to migrate, e.g. Contact, Calendar, etc. If not specified, all item types will be migrated.
    $migration = Add-MW_MailboxMigration -Ticket $mwTicket -ConnectorId $connector.Id -MailboxId $_.Id -UserId $mwTicket.UserId -Type Full -ItemTypes Contact   
}