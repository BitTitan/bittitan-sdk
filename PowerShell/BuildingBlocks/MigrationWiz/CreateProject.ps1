# This script illustrates how to create a project in MigrationWiz

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
$ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan
$mwTicket = Get-MW_Ticket -Credentials $creds

# Retrieve a customer, in order to create the project under that customer 
$customer = Get-BT_Customer -Ticket $ticket -CompanyName "Default"

# In this example, we create a mailbox project from Exchange Server to O365
# Set up export and import configurations
# You can also choose to provide admin credentials and set UseAdministrativeCredentials to true, then you do not need to provide usernames and passwords in mailbox creation.
$exportConfiguration = New-Object -TypeName MigrationProxy.WebApi.ExchangeConfiguration -Property @{
        "UseAdministrativeCredentials" = $false;
        "Url" = "https://exchange-server-url.com";
        "ExchangeItemType" = "Mail";
}
$importConfiguration = New-Object -TypeName MigrationProxy.WebApi.ExchangeConfiguration -Property @{
        "UseAdministrativeCredentials" = $false;
        "ExchangeItemType" = "Mail";
}

# Create the project
$connector = Add-MW_MailboxConnector -Ticket $mwTicket -ProjectType Mailbox -ExportType ExchangeServer -ImportType Office365 `
    -Name "TestProject" -UserId $mwTicket.UserId -OrganizationId $customer.OrganizationId `
    -ExportConfiguration $exportConfiguration -ImportConfiguration $importConfiguration