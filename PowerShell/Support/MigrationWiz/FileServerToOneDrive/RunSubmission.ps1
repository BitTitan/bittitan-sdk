Import-Module "C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll"

$cred = Get-Credential

$t = Get-MW_Ticket -Credentials $cred 

#$connectors = Get-MW_MailboxConnector -Ticket $t -FilterBy_String_Name "%_PS"
$connectors = Get-MW_MailboxConnector -Ticket $t |where-object {$_.Name -like "*_PS"}

foreach ($connector in $connectors)
{

    $PageOffSet = 0
    $PageSize = 50

    $rows = Get-MW_Mailbox -Ticket $t -FilterBy_Guid_ConnectorId $connector.Id -PageOffset $PageOffSet -PageSize $PageSize

    while ($rows)
    {
        foreach ($row in $rows)
        {
            $lastMigrationAttempt = Get-MW_MailboxMigration -ticket $t -FilterBy_Guid_MailboxId $row.Id -SortBy_CreateDate_Descending | Select-Object -Property MailboxId, CompleteDate, Status, Type, ItemTypes | Select -First 1
    
            if ($lastMigrationAttempt.Status -eq "Completed")
            {
                if ($lastMigrationAttempt.Type -eq "Verification")
                {
                    $submission = Add-MW_MailboxMigration -Ticket $t -MailboxId $row.Id -Type Full -Status Submitted -ConnectorId $connector.Id -UserId $t.UserId -Priority 1 -MaximumItemsPerFolder 2147483647 
                }  
                if ($lastMigrationAttempt.Type -eq "Full")
                {
                    $Usermail = $row.ImportEmailAddress
                    Write-host $usermail
                    $msg = "A full migration was already executed for user $Usermail. An additional one will be processed now."
                    Write-host $msg -ForegroundColor Red
                    $submission = Add-MW_MailboxMigration -Ticket $t -MailboxId $row.Id -Type Full -Status Submitted -ConnectorId $connector.Id -UserId $t.UserId -Priority 1 -MaximumItemsPerFolder 2147483647 
                }  
            }
            if ($lastMigrationAttempt.Status -eq "Failed")
            {
                $Usermail = $row.ImportEmailAddress
                Write-host $usermail
                $msg = "A full migration attempt for user $Usermail as failed. A new attempt will be processed now."
                Write-host $msg -ForegroundColor Red
                $submission = Add-MW_MailboxMigration -Ticket $t -MailboxId $row.Id -Type Full -Status Submitted -ConnectorId $connector.Id -UserId $t.UserId -Priority 1 -MaximumItemsPerFolder 2147483647 
            }
        }

        $PageOffSet += $PageSize

        $rows = Get-MW_Mailbox -Ticket $t -FilterBy_Guid_ConnectorId $connector.Id -PageOffset $PageOffSet -PageSize $PageSize
    }
}