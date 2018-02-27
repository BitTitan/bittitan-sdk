<#
.NOTES
    Company:		BitTitan, Inc.
    Title:			BitTitanManagement.PS1
    Author:			SUPPORT@BITTITAN.COM
    Requirements: 
    
    Version:		1.00
    Date:			DECEMBER 1, 2016

    Windows Version:	WINDOWS 10 ENTERPRISE

    Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

    Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
    
.SYNOPSIS
    Provides a set of console tools for MigrationWiz.

.DESCRIPTION 	
    This script provides a set of tools to manage migration projects, local exchange servers, G Suite and Office365. 

.INPUTS	

.EXAMPLE
    .\BitTitanManagement.ps1
    Launches the MigrationWiz console tool.
#>

$connectorPageSize = 100
$mailboxPageSize = 100
$precision = "0.00"
$ticketGracePeriodMinutes = 60
$maximumAttempts = 5
$retrySleepSeconds = 30
$stopSleepSeconds = 30
$replicationSleepSeconds = 5
$debug = $false
$o365PowerShellDownloadLink = "http://technet.microsoft.com/en-us/library/hh974317.aspx"
$migrationAdminName = "MigrationWiz"
$exchangeContactOuName = "External Forwards"

$migrationWizCreds = $null
$migrationWizTicket = $null
$googleCreds = $null
$o365Creds = $null
$environment = $null
$allExchangeServerPublicFoldersFileName = $null
$allExchangeServerPublicFoldersPerMailboxFileName = $null

######################################################################################################################################################
# MigrationWiz commands
######################################################################################################################################################

function MWHelper-ChooseEnvironment()
{
    # Set environment
    if($script:environment -eq $null)
    {
        $script:environment = Read-Host -Prompt "Select environment. Options include BT (default) or China. Press <Enter> to select default"
    }

    if(!$script:environment)
    {
        $script:environment = "BT"
    }

    return $script:environment
}

function MW-GetTicket([System.Management.Automation.PSCredential]$creds)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $ticket = Get-MW_Ticket -Credentials $creds -Environment $script:environment
            return $ticket
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-GetMailboxStats([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $stats = Get-MW_MailboxStat -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -FilterBy_Guid_MailboxId $mailbox.Id
            return $stats
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage  -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-GetMailboxMigrationLatest([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $latest = Get-MW_MailboxMigration -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -FilterBy_Guid_MailboxId $mailbox.Id -SortBy_CompleteDate_Descending -PageSize 1 -PageOffset 0
            return $latest
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage  -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-GetMailboxErrors([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $errors = Get-MW_MailboxError -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -FilterBy_Guid_MailboxId $mailbox.Id
            return $errors
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-GetMailboxHistory([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $history = @(Get-MW_MailboxMigration -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -FilterBy_Guid_MailboxId $mailbox.Id -SortBy_CreateDate_Ascending)
            return ,$history
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-SubmitMigration([MigrationProxy.WebApi.Mailbox]$mailbox, [MigrationProxy.WebApi.MailboxQueueTypes]$licenseType)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $ticket = MWHelper-GetTicket
            $migration = Add-MW_MailboxMigration -Ticket $ticket -Environment (MWHelper-ChooseEnvironment) -MailboxId $mailbox.Id -ConnectorId $mailbox.ConnectorId -Type $licenseType -UserId $ticket.UserId -Priority 1 -Status Submitted
            return
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage  -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-StopMigration([MigrationProxy.WebApi.MailboxMigration]$migration)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $migration = Set-MW_MailboxMigration -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -mailboxmigration $migration -Status Stopping
            return
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage  -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-GetMailboxConnectors([int]$connectorOffSet, [int]$connectorPageSize)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $connectors = @(Get-MW_MailboxConnector -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -PageOffset $connectorOffSet -PageSize $connectorPageSize)
            return ,$connectors
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage  -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-GetMailboxes([MigrationProxy.WebApi.MailboxConnector]$connector, [int]$mailboxOffSet, [int]$mailboxPageSize)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $mailboxes = @(Get-MW_Mailbox -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -FilterBy_Guid_ConnectorId $connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize)
            return ,$mailboxes
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-RemoveMailboxConnector([MigrationProxy.WebApi.MailboxConnector]$connector)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            Remove-MW_MailboxConnector -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -Id $connector.Id -Force:$true
            return
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage  -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-SetMailboxConnector([MigrationProxy.WebApi.MailboxConnector]$connector)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $connector = Set-MW_MailboxConnector -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -mailboxconnector $connector -ItemStartDate $connector.ItemStartDate -ItemEndDate $connector.ItemEndDate -FolderFilter $connector.FolderFilter -Flags $connector.Flags -DisabledMailboxItemTypes $connector.DisabledMailboxItemTypes -MaximumSimultaneousMigrations $connector.MaximumSimultaneousMigrations -MaximumItemFailures $connector.MaximumItemFailures -PurgePeriod $connector.PurgePeriod -AdvancedOptions $connector.AdvancedOptions
            return
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage  -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-AddMailbox([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $attempts = 0
    while ($true)
    {
        try
        {
            $attempts++
            $mailbox = Add-MW_Mailbox -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -ConnectorId $mailbox.ConnectorId -ExportEmailAddress $mailbox.ExportEmailAddress -ExportPassword $mailbox.ExportPassword -ExportUserName $mailbox.ExportUserName -Flags $mailbox.Flags -FolderFilter $mailbox.FolderFilter -PublicFolderPath $mailbox.PublicFolderPath -ImportEmailAddress $mailbox.ImportEmailAddress -ImportPassword $mailbox.ImportPassword -ImportUserName $mailbox.ImportUserName -AdvancedOptions $mailbox.AdvancedOptions
            return
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

function MW-SetMailbox([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $attempts = 0
    while($true)
    {
        try
        {
            $attempts++
            $mailbox = Set-MW_Mailbox -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -mailbox $mailbox -FolderFilter $mailbox.FolderFilter -Flags $mailbox.Flags -DisabledMailboxItemTypes $mailbox.DisabledMailboxItemTypes -AdvancedOptions $mailbox.AdvancedOptions
            return
        }
        catch
        {
            if($attempts -ge $script:maximumAttempts)
            {
                throw
            }

            Helper-DisplayRetryMessage -attempts $attempts -errorMessage $_.ToString()
            Start-Sleep -Seconds $script:retrySleepSeconds
        }
    }
}

######################################################################################################################################################
# MigrationWiz helper functions
######################################################################################################################################################

function MWHelper-GetTicket()
{
    if($script:migrationWizCreds -eq $null)
    {
        # prompt for credentials
        $script:migrationWizCreds = $host.ui.PromptForCredential("BitTitan Credentials", "Enter your BitTitan user name and password", "", "")
    }

    if($script:migrationWizCreds -ne $null)
    {
        # get new ticket if we don't already have one or it's expired
        if(($script:migrationWizTicket -eq $null) -or ($script:migrationWizTicket.ExpirationDate.AddMinutes(-1*$ticketGracePeriodMinutes) -lt (Get-Date).ToUniversalTime()))
        {
            # get new ticket
            $script:migrationWizTicket = MW-GetTicket -creds $script:migrationWizCreds

            if($script:migrationWizTicket -ne $null)
            {
                Helper-WriteDebug -line ("MigrationWiz ticket will expire on " + $script:migrationWizTicket.ExpirationDate.ToLocalTime().ToString())
            }
        }
    }

    return $script:migrationWizTicket
}

function MWHelper-GetConnectors()
{
    $connectorOffSet = 0
    $connectors = $null

    Write-Host
    Write-Host -Object  "Retrieving mailbox connectors ..."

    do
    {
        $connectorsPage = MW-GetMailboxConnectors -connectorOffSet $connectorOffSet -connectorPageSize $script:connectorPageSize
        if($connectorsPage)
        {
            $connectors += @($connectorsPage)
            foreach($connector in $connectorsPage)
            {
                Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }
    }
    while($connectorsPage)

    if($connectors -ne $null -and $connectors.Length -ge 1)
    {
        Write-Host -Object  ($connectors.Length.ToString() + " mailbox connector(s) found")
    }
    else
    {
        Write-Host -Object  "No mailbox connectors found" -ForegroundColor Yellow
    }


    return ,$connectors
}

function MWHelper-GetMailboxes([MigrationProxy.WebApi.MailboxConnector]$connector)
{
    $mailboxOffSet = 0
    $mailboxes = $null

    Write-Host
    Write-Host -Object  ("Retrieving mailboxes for " + $connector.Name)

    do
    {
        $mailboxesPage = MW-GetMailboxes -connector $connector -mailboxOffSet $mailboxOffSet -mailboxPageSize $script:mailboxPageSize
        if($mailboxesPage)
        {
            $mailboxes += @($mailboxesPage)
            foreach($mailbox in $mailboxesPage)
            {
                Write-Progress -Activity ("Retrieving mailboxes for " + $connector.Name + " (" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress
            }

            $mailboxOffSet += $mailboxPageSize
        }
    }
    while($mailboxesPage)

    if($mailboxes -ne $null -and $mailboxes.Length -ge 1)
    {
        Write-Host -Object  ($mailboxes.Length.ToString() + " mailbox(es) found")
    }
    else
    {
        Write-Host -Object  "No mailboxes found" -ForegroundColor Yellow
    }

    return ,$mailboxes
}

function MWHelper-GetMailboxStatistics([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $folderSuccessSize = 0
    $calendarSuccessSize = 0
    $contactSuccessSize = 0
    $mailSuccessSize = 0
    $taskSuccessSize = 0
    $noteSuccessSize = 0
    $journalSuccessSize = 0
    $rulesSuccessSize = 0
    $totalSuccessSize = 0

    $folderSuccessCount = 0
    $calendarSuccessCount = 0
    $contactSuccessCount = 0
    $mailSuccessCount = 0
    $taskSuccessCount = 0
    $noteSuccessCount = 0
    $journalSuccessCount = 0
    $rulesSuccessCount = 0
    $totalSuccessCount = 0

    $folderErrorSize = 0
    $calendarErrorSize = 0
    $contactErrorSize = 0
    $mailErrorSize = 0
    $taskErrorSize = 0
    $noteErrorSize = 0
    $journalErrorSize = 0
    $rulesErrorSize = 0
    $totalErrorSize = 0

    $folderErrorCount = 0
    $calendarErrorCount = 0
    $contactErrorCount = 0
    $mailErrorCount = 0
    $taskErrorCount = 0
    $noteErrorCount = 0
    $journalErrorCount = 0
    $rulesErrorCount = 0
    $totalErrorCount = 0

    $totalExportActiveDuration = 0
    $totalExportPassiveDuration = 0
    $totalImportActiveDuration = 0
    $totalImportPassiveDuration = 0

    $totalExportSpeed = 0
    $totalExportCount = 0

    $totalImportSpeed = 0
    $totalImportCount = 0

    $stats = MW-GetMailboxStats -mailbox $mailbox

    $Calendar = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Calendar)
    $Contact = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Contact)
    $Mail = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Mail)
    $Journal = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Journal)
    $Note = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Note)
    $Task = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Task)
    $Folder = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Folder)
    $Rule = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Rule)

    if($stats -ne $null)
    {
        foreach($info in $stats.MigrationStatsInfos)
        {
            switch ([int]$info.ItemType)
            {
                $Folder
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $folderSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $folderSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $folderErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $folderErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Calendar
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $calendarSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $calendarSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $calendarErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $calendarErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Contact
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $contactSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $contactSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $contactErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $contactErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Mail
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $mailSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $mailSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $mailErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $mailErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Task
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $taskSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $taskSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $taskErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $taskErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Note
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $noteSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $noteSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $noteErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $noteErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Journal
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $journalSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $journalSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $journalErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $journalErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                $Rule
                {
                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import)
                    {
                        $ruleSuccessSize = $info.MigrationStats.SuccessSize + $info.MigrationStats.SuccessSizeTotal
                        $ruleSuccessCount = $info.MigrationStats.SuccessCount + $info.MigrationStats.SuccessCountTotal
                        $ruleErrorSize = $info.MigrationStats.ErrorSize + $info.MigrationStats.ErrorSizeTotal
                        $ruleErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                    }
                    break
                }

                default {break}
            }
        }

        $totalSuccessSize = $folderSuccessSize + $calendarSuccessSize + $contactSuccessSize + $mailSuccessSize + $taskSuccessSize + $noteSuccessSize + $journalSuccessSize + $rulesSuccessSize
        $totalSuccessCount = $folderSuccessCount + $calendarSuccessCount + $contactSuccessCount + $mailSuccessCount + $taskSuccessCount + $noteSuccessCount + $journalSuccessCount + $rulesSuccessCount
        $totalErrorSize = $folderErrorSize + $calendarErrorSize + $contactErrorSize + $mailErrorSize + $taskErrorSize + $noteErrorSize + $journalErrorSize + $rulesErrorSize
        $totalErrorCount = $folderErrorCount + $calendarErrorCount + $contactErrorCount + $mailErrorCount + $taskErrorCount + $noteErrorCount + $journalErrorCount + $rulesErrorCount

        $totalExportActiveDuration = ($stats.ExportDuration - $stats.WaitExportDuration) / 1000 / 60
        $totalExportPassiveDuration = $stats.WaitExportDuration / 1000 / 60
        $totalImportActiveDuration = ($stats.ImportDuration - $stats.WaitImportDuration) / 1000 / 60
        $totalImportPassiveDuration = $stats.WaitImportDuration / 1000 / 60

        if($totalSuccessSize -gt 0 -and $totalExportActiveDuration -gt 0)
        {
            $totalExportSpeed = $totalSuccessSize / 1024 / 1024 / $totalExportActiveDuration * 60
            $totalExportCount = $totalSuccessCount / $totalExportActiveDuration * 60
        }

        if($totalSuccessSize -gt 0 -and $totalImportActiveDuration -gt 0)
        {
            $totalImportSpeed = $totalSuccessSize / 1024 / 1024 / $totalImportActiveDuration * 60
            $totalImportCount = $totalSuccessCount / $totalImportActiveDuration * 60
        }
    }

    return @(($stats -ne $null),$folderSuccessSize,$calendarSuccessSize,$contactSuccessSize,$mailSuccessSize,$taskSuccessSize,$noteSuccessSize,$journalSuccessSize,$totalSuccessSize,$folderSuccessCount,$calendarSuccessCount,$contactSuccessCount,$mailSuccessCount,$taskSuccessCount,$noteSuccessCount,$journalSuccessCount,$totalSuccessCount,$folderErrorSize,$calendarErrorSize,$contactErrorSize,$mailErrorSize,$taskErrorSize,$noteErrorSize,$journalErrorSize,$totalErrorSize,$folderErrorCount,$calendarErrorCount,$contactErrorCount,$mailErrorCount,$taskErrorCount,$noteErrorCount,$journalErrorCount,$totalErrorCount,$totalExportActiveDuration,$totalExportPassiveDuration,$totalImportActiveDuration,$totalImportPassiveDuration,$totalExportSpeed,$totalExportCount,$totalImportSpeed,$totalImportCount)
}

function MWHelper-CreateMailboxes($connector, [string]$pfListCsvPath)
{
    # Split Public Folder into multiple migrations
    $migrations = Split-PublicFolderMigrations -PublicFolderCsvPath $pfListCsvPath

    # Create mailboxes from the migration list
    Write-Host -Object "Creating project items inside '$($connector.Name)' project"
    foreach($migration in $migrations)
    {
        $mailbox = Add-MW_Mailbox -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -ConnectorId $connector.Id -FolderFilter $migration.Value -PublicFolderPath $migration.Key
        Write-Host -Object "  Created item with public folder path '$($migration.Key)'"
    }
    Write-Host -Object "  Completed adding $($migrations.Length) project items to the '$($connector.Name)' project)"
}

######################################################################################################################################################
# Google helper functions
######################################################################################################################################################

function GoogleHelper-GetDomainUsers([string]$domainName)
{
    Write-Host -Object "poop"
    $users = Get-MigrationWizGoogleUserAccounts -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -DomainName $domainName
    return $users
}

function GoogleHelper-SetMailboxForward([string]$emailAddress, [string]$targetAddress)
{
    $forward = Set-MigrationWizGoogleEmailForward -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -EmailAddress $emailAddress -Enable $true -ForwardTo $targetAddress -Action ([BitTitan.Powershell.Core.GoogleEmailForwardAction]::Archive)
}

function GoogleHelper-RemoveMailboxForward([string]$emailAddress)
{
    $forward = Set-MigrationWizGoogleEmailForward -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment) -EmailAddress $emailAddress -Enable $false -ForwardTo $emailAddress -Action ([BitTitan.Powershell.Core.GoogleEmailForwardAction]::Null)
}

function GoogleHelper-RevokeOAuth2AccessToken()
{
    $forward = Remove-MigrationWizGoogleOAuth2AccessToken -Ticket (MWHelper-GetTicket) -Environment (MWHelper-ChooseEnvironment)
}

######################################################################################################################################################
# Office 365 helper functions
######################################################################################################################################################

function Office365Helper-GetCredentials()
{
    if($script:o365Creds -eq $null)
    {
        # prompt for credentials
        $script:o365Creds = $host.ui.PromptForCredential("Office 365 Credentials", "Enter your Office 365 administrative user name and password", "", "")
    }

    return $script:o365Creds
}

function Office365Helper-PromptSku()
{
    Write-Host
    Write-Host -Object  "Select an Office 365 license to assign:" -ForegroundColor Yellow
    Write-Host

    $skus = @(Get-MsolAccountSku)
    if($skus -ne $null)
    {
        for ($i=0; $i -lt $skus.Length; $i++)
        {
            $sku = $skus[$i]
            Write-Host -Object  ("$i - " + $sku.SkuPartNumber + " (" + $sku.ConsumedUnits + "/" + $sku.ActiveUnits + ")")
        }
        Write-Host

        do
        {
            $result = Read-Host -Prompt ("Select 0-" + ($skus.Length-1))
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $skus.Length))
            {
                return $skus[$result]
            }
        }
        while($true)
    }

    return $null
}

function Office365Helper-WaitForReplication([string]$userPrincipalName)
{
    while($true)
    {
        $user = Get-MsolUser -UserPrincipalName $userPrincipalName -ErrorAction SilentlyContinue
        if($user -ne $null)
        {
            return $user
        }

        Write-Host -Object  ("Waiting for Office 365 replication. Retry in $script:replicationSleepSeconds seconds.")
        Start-Sleep -Seconds $script:replicationSleepSeconds
    }
}

function Office365Helper-ConnectRemotePowerShell()
{
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential (Office365Helper-GetCredentials) -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
    $import = Import-PSSession -Session $session -AllowClobber -WarningAction SilentlyContinue
}

######################################################################################################################################################
# Exchange Server helper functions
######################################################################################################################################################

function ExchangeServerHelper-GetCredentials()
{
    if($script:exchangeServerCreds -eq $null)
    {
        # prompt for credentials
        $script:exchangeServerCreds = $host.ui.PromptForCredential("Exchange Server Credentials", "Enter your Exchange Server administrative user name and password", "", "")
    }

    return $script:exchangeServerCreds
}

function ExchangeServerHelper-ConnectPowerShell()
{
    # Get the fully qualified domain name of the exchange server
    $fqdn = (Helper-PromptString -prompt "Exchange Server Domain Name [$env:computername.$env:userdnsdomain]" -allowEmpty $true)
    if($fqdn.Length -le 1)
    {
        $fqdn = "$env:computername.$env:userdnsdomain"
    }

    # Create and import the PowerShell session
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://"$fqdn"/PowerShell/ -Authentication Kerberos -Credential (ExchangeServerHelper-GetCredentials)
    $import = Import-PSSession -Session $session -AllowClobber -WarningAction SilentlyContinue
}

function ExchangeServerHelper-SetPublicFolderFileName([string]$fileName)
{
    if ($fileName -eq $null)
    {
        throw "The file name for all public folders cannot be null"
    }

    $script:allExchangeServerPublicFoldersFileName = $pfFileName
}

function ExchangeServerHelper-SetPublicFolderPerMailboxFileName([string]$fileName)
{
    if ($fileName -eq $null)
    {
        throw "The file name of the public folders per mailbox cannot be null"
    }

    $script:allExchangeServerPublicFoldersPerMailboxFileName = $fileName
}

function ExchangeServerHelper-AdSearch([string]$searchFilter, [string[]]$searchProperties)
{
    Helper-WriteDebug -line ("AD Search Filter = $searchFilter")
    if($searchProperties -ne $null)
    {
        Helper-WriteDebug -line ("AD Load Properties = " + ($searchProperties -join ","))
    }

    $rootDse = [ADSI]"LDAP://RootDSE"
    $defaultNamingContext = $rootDse.defaultNamingContext
    $dc = $rootDse.dnsHostName
    $searcher = New-Object -TypeName DirectoryServices.DirectorySearcher
    $searcher.CacheResults = $false
    $searcher.Filter = $searchFilter
    $searcher.PageSize = 100;
    if($searchProperties -ne $null -and $searchProperties.Length -ge 1)
    {
        $result = $searcher.PropertiesToLoad.AddRange($searchProperties)
    }
    $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
    $searcher.SearchRoot = "LDAP://$dc/$defaultNamingContext"
    $objects = @($searcher.FindAll())

    if($objects -ne $null)
    {
        Helper-WriteDebug -line ("AD Search Results = " + $objects.Length)
    }
    else
    {
        Helper-WriteDebug -line ("AD Search Results = NULL")
    }

    return $objects
}

function ExchangeServerHelper-CreateOu([string]$name)
{
    Helper-WriteDebug -line ("OU Name = $name")

    $rootDse = [ADSI]"LDAP://RootDSE"
    $defaultNamingContext = $rootDse.defaultNamingContext
    $dc = $rootDse.dnsHostName
    $ouPath = "LDAP://$dc/OU=$name,$defaultNamingContext"

    Helper-WriteDebug -line ("OU Path = $ouPath")

    if([System.DirectoryServices.DirectoryEntry]::Exists($ouPath))
    {
        Helper-WriteDebug -line ("OU already exists")
        $ou = [ADSI]$ouPath
    }
    else
    {
        Helper-WriteDebug -line ("Creating new OU")

        $domain = [ADSI]"LDAP://$dc/$defaultNamingContext"
        $ou = $domain.Children.Add("OU=$name", "organizationalUnit")
        $ou.CommitChanges()
    }

    return $ou
}

function ExchangeServerHelper-CopyProperty([ADSI]$source, [ADSI]$destination, [string]$property)
{
    if($source.Properties[$property] -ne $null)
    {
        if($source.Properties[$property][0] -ne "")
        {
            $destination.Put($property, $source.Properties[$property][0])
        }
    }
}

function ExchangeServerHelper-CreateContactForward([ADSI]$user, [string]$targetAddress, [bool]$setForward)
{
    Helper-WriteDebug -line ("Creating contact forward for $targetAddress")

    $rand = New-Object -TypeName System.Random
    $id = $rand.Next().ToString()
    $proxyAddresses = @()

    if($setForward)
    {
        $proxyAddresses += @("smtp:" + $targetAddress)
        foreach($proxyAddress in $user.Properties["proxyAddresses"])
        {
            Helper-WriteDebug -line ("Proxy Address: " + $proxyAddress)
            $proxyAddresses += @($proxyAddress)
        }
    }
    else
    {
        $proxyAddresses += @("SMTP:" + $targetAddress)
    }

    $ou = (ExchangeServerHelper-CreateOu -name $exchangeContactOuName)
    $name = (Helper-ReplaceNonAlphaNumeric -s $user.Properties["name"][0] -replacement "_")
    if($name.Length -gt 20) { $name = $name.SubString(0, 20) }
    $name += $id

    Helper-WriteDebug -line ("Contact forward CN is $name")

    $contact = $ou.Children.Add("CN=" + $name, "contact")
    $contact.Put("mail", $targetAddress)
    $contact.Put("msExchHideFromAddressLists", $true)
    $contact.Put("targetAddress", "SMTP:$targetAddress")
    $contact.Put("proxyAddresses", $proxyAddresses)

    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "company"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "department"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "displayName"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "givenName"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "initials"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "l"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "mailNickname"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "physicalDeliveryOfficeName"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "postalCode"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "sn"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "st"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "streetAddress"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "telephoneNumber"
    ExchangeServerHelper-CopyProperty -source $user -destination $contact -property "title"

    $contact.CommitChanges()

    if($setForward)
    {
        $dn = $contact.Properties["distinguishedName"][0]
        Helper-WriteDebug -line ("Contact forward DN is $dn")

        $newEmailAddress = $user.Properties["mail"][0]
        $newEmailAddress = $newEmailAddress.Replace("@", "-MigrationWiz@")

        $user.Put("altRecipient", $dn)
        $user.Put("mail", $newEmailAddress)
        $user.Put("proxyAddresses", ("SMTP:" + $newEmailAddress))
        #$user.Put("msExchHideFromAddressLists", $true)
        $user.CommitChanges()
    }
}

######################################################################################################################################################
# Helper functions
######################################################################################################################################################

function Helper-GenerateRandomTempFilename([string]$identifier)
{
    $filename = $env:temp + "\MigrationWiz-"
    if($identifier -ne $null -and $identifier.Length -ge 1)
    {
        $filename += $identifier + "-"
    }
    $filename += (Get-Date).ToString("yyyyMMddHHmmss")
    $filename += ".csv"

    return $filename
}

function Helper-GenerateRandomTempProjectName([string]$identifier)
{
    $projectName = "MigrationWiz-"
    if($identifier -ne $null -and $identifier.Length -ge 1)
    {
        $projectName += $identifier + "-"
    }
    $projectName += (Get-Date).ToString("yyyyMMddHHmmss")

    return $projectName
}

function Helper-DisplayRetryMessage([int]$attempts, [string]$errorMessage)
{
    Write-Progress -Activity ("Error encountered while communicating with MigrationWiz ... retry in " + $retrySleepSeconds + " seconds (" + $attempts + "/" + $script:maximumAttempts + " attempts made)") -Status $errorMessage
}

function Helper-IncreaseWindowSize([int]$width, [int]$height)
{
    # Returns if it is window size is null; this happens when running in PowerShell ISE
    if($host.ui.rawui.WindowSize -eq $null){
        Return
    }

    $maxWindowWidth = $host.ui.rawui.MaxPhysicalWindowSize.Width
    $maxWindowHeight = $host.ui.rawui.MaxPhysicalWindowSize.Height

    $curWindowWidth = $host.ui.rawui.WindowSize.Width
    $curWindowHeight = $host.ui.rawui.WindowSize.Height

    $newWindowWidth = [math]::min($width, $maxWindowWidth)
    $newWindowHeight = [math]::min($height, $maxWindowHeight)

    if($curWindowWidth -lt $newWindowWidth)
    {
        $bufferSize = $host.ui.rawui.BufferSize;
        $bufferSize.width = $newWindowWidth
        $host.ui.rawui.BufferSize = $bufferSize

        $windowSize = $host.ui.rawui.WindowSize;
        $windowSize.width = $newWindowWidth
        $host.ui.rawui.WindowSize = $windowSize
    }

    if($curWindowHeight -lt $newWindowHeight)
    {
        $windowSize = $host.ui.rawui.WindowSize;
        $windowSize.height = $newWindowHeight
        $host.ui.rawui.WindowSize = $windowSize
    }
}

function Helper-LoadModule([string]$name, [string]$filename, [bool]$fatal)
{
    $loaded = $false

    if((Get-Module -Name $name) -eq $null)
    {
        if(Test-Path -Path $filename)
        {
            Import-Module -Name $filename

            Helper-WriteDebug -line ("The PowerShell module " + $name + " was successfully loaded")
            $loaded = $true
        }
        else
        {
            Helper-WriteDebug -line ("The PowerShell module file " + $filename + " was not found")
        }
    }
    else
    {
        Helper-WriteDebug -line ("The PowerShell module " + $name + " was already loaded")
        $loaded = $true
    }

    if($fatal)
    {
        throw ("Could not load module " + $filename)
    }

    return $loaded
}

function Helper-WriteDebug([string]$line)
{
    if($debug)
    {
        Write-Host -Object  ("DEBUG: $line")
    }
}

function Helper-LoadMigrationWizModule()
{
    if (((Get-Module -Name "BitTitanPowerShell") -ne $null) -or ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null))
    {
        return;
    }

    $currentPath = Split-Path -parent -Path $script:MyInvocation.MyCommand.Definition
    $moduleFilename = "$currentPath\BitTitanPowerShell.dll"
    if((Helper-LoadModule -name "BitTitanPowerShell" -filename $moduleFilename -fatal $false) -eq $false)
    {
        $moduleFilename = "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll"
        if((Helper-LoadModule -name "BitTitanPowerShell" -filename $moduleFilename -fatal $false) -eq $false)
        {
            $moduleFilename = "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll"
            if((Helper-LoadModule -name "BitTitanPowerShell" -filename $moduleFilename -fatal $false) -eq $false)
            {
                throw "Could not load BitTitan PowerShell module"
            }
        }
    }
}

function Helper-LoadOffice365Module()
{
    if((Get-Module -Name MSOnline) -eq $null)
    {
        Helper-WriteDebug -line ("Office 365 PowerShell module was not loaded")

        Import-Module -Name MSOnline -ErrorAction SilentlyContinue
        if((Get-Module -Name MSOnline) -eq $null)
        {
            Helper-WriteDebug -line ("Office 365 PowerShell module was not found")
            Start-Process -FilePath $o365PowerShellDownloadLink
            throw ("The Office 365 PowerShell module was not found.  Download and install the Microsoft Online Services Sign-In Assistant and the Microsoft Online Services Module for Windows PowerShell from " + $o365PowerShellDownloadLink)
        }
        else
        {
            Helper-WriteDebug -line ("Office 365 PowerShell module successfully loaded")
        }
    }
    else
    {
        Helper-WriteDebug -line ("Office 365 PowerShell module is already loaded")
    }
}

function Helper-PromptConfirmation([string]$prompt)
{
    while($true)
    {
        $confirm = Read-Host -Prompt ($prompt + " [Y]es or [N]o")

        if($confirm -eq "Y")
        {
            return $true
        }

        if($confirm -eq "N")
        {
            return $false
        }
    }
}

function Helper-PromptString([string]$prompt, [bool]$allowEmpty)
{
    while($true)
    {
        $value = Read-Host -Prompt ($prompt)

        if($value.Length -ge 1)
        {
            return $value
        }

        if($allowEmpty -and $value.Length -eq 0)
        {
            return ""
        }
    }
}

function Helper-StringInArray([string]$toFind, [string[]]$stringArray)
{
    foreach($s in $stringArray)
    {
        if($toFind.ToLower() -eq $s.ToLower())
        {
            return $true
        }
    }

    return $false
}

function Helper-GeneratePassword()
{
    $upperCaseChars = "ABCDEFGHIJKLMNPQRSTUVWXYZ"
    $lowerCaseChars = "abcdefghijkmnopqrstuvwxyz"
    $numericChars = "23456789"
    $symbolChars = "-=!@#$%^&*()_+"

    $password = ""

    $rand = New-Object -TypeName System.Random
    1..1 | ForEach-Object -Process { $password = $password + $upperCaseChars[$rand.next(0,$upperCaseChars.Length-1)] }
    1..7 | ForEach-Object -Process { $password = $password + $lowerCaseChars[$rand.next(0,$lowerCaseChars.Length-1)] }
    1..3 | ForEach-Object -Process { $password = $password + $numericChars[$rand.next(0,$numericChars.Length-1)] }
    1..1 | ForEach-Object -Process { $password = $password + $symbolChars[$rand.next(0,$symbolChars.Length-1)] }

    return $password
}

function Helper-ReplaceNonAlphaNumeric([string]$s, [string]$replacement)
{
    $result = ""

    foreach($c in $s.GetEnumerator())
    {
        if ([System.Char]::IsLetterOrDigit($c))
        {
            $result += $c
        }
        else
        {
            $result += $replacement
        }
    }

    return $result
}

######################################################################################################################################################
# Main menu
######################################################################################################################################################

function Menu-Banner()
{
    Write-Host
    Write-Host -Object  "Main Menu:" -ForegroundColor Yellow
    Write-Host

    return $null
}

function Menu-MainPrompt()
{
    Write-Host
    Write-Host -Object  "0 - Manage migration projects"
    Write-Host -Object  "1 - Manage G Suite"
    Write-Host -Object  "2 - Manage Office 365"
    Write-Host -Object  "3 - Manage local Exchange Server"
    Write-Host -Object  "x - Exit"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0-3 or x"
        if($result -eq "x")
        {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 3))
        {
            return [int]$result
        }
    }

    return $null
}

function Menu-MainLoop()
{
    # keep looping until specified to exit
    do
    {
        try
        {
            $action = Menu-MainPrompt
            if($action -ne $null)
            {
                switch($action)
                {
                    0 # Manage migration projects
                    {
                        Menu-MigrationWizConnectorListLoop
                    }

                    1 # Manage G Suite
                    {
                        $script:googleCreds = $null
                        Menu-GoogleLoop
                    }

                    2 # Manage Office 365
                    {
                        $script:o365Creds = $null
                        Menu-Office365Loop
                    }

                    3 # Manage local Exchange Server
                    {
                        Menu-ExchangeServerLoop
                    }
                }
            }
            else
            {
                return
            }
        }
        catch
        {
            Write-Host
            Write-Host -Object  $_.ToString() -ForegroundColor Red
            Write-Host
            Write-Output -InputObject $_
        }
    }
    while($true)
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors
######################################################################################################################################################

function Menu-MigrationWizConnectorListPrompt([MigrationProxy.WebApi.MailboxConnector[]]$connectors)
{
    if($connectors -ne $null)
    {
        Write-Host
        Write-Host -Object  "Select a mailbox connector:" -ForegroundColor Yellow
        Write-Host

        for ($i=0; $i -lt $connectors.Length; $i++)
        {
            $connector = $connectors[$i]
            Write-Host -Object $i,"-",$connector.Name
        }
        Write-Host -Object "x - Back"
        Write-Host

        do
        {
            $result = Read-Host -Prompt ("Select 0-" + ($connectors.Length-1) + " or x")
            if($result -eq "x")
            {
                return $null
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $connectors.Length))
            {
                return $connectors[$result]
            }
        }
        while($true)
    }

    return $null
}

function Menu-MigrationWizConnectorListLoop()
{
    if(MWHelper-GetTicket)
    {
        # keep looping until specified to exit
        do
        {
            $connectors = MWHelper-GetConnectors

            if($connectors -ne $null)
            {
                $connector = Menu-MigrationWizConnectorListPrompt -connectors $connectors
                if($connector -ne $null)
                {
                    Menu-MigrationWizConnectorTaskLoop -connector $connector
                }
                else
                {
                    return
                }
            }
            else
            {
                return
            }
        }
        while($true)
    }
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Mailbox task menu
######################################################################################################################################################

function Menu-MigrationWizConnectorTaskPrompt([MigrationProxy.WebApi.MailboxConnector]$connector)
{
    Write-Host
    Write-Host -Object  ("Select a task to perform on " + $connector.Name + ":") -ForegroundColor Yellow
    Write-Host
    Write-Host -Object  "0 - Display mailbox migration status"
    Write-Host -Object  "1 - Export statistics and errors to CSV"
    Write-Host -Object  "2 - Submit mailboxes for migration"
    Write-Host -Object  "3 - Stop all mailbox migrations"
    Write-Host -Object  "4 - Delete mailbox connector"
    Write-Host -Object  "5 - Configure mailbox connector"
    Write-Host -Object  "6 - Reset configuration on all mailboxes to default"
    Write-Host -Object  "7 - Configure public folder to shared mailbox connector"
    Write-Host -Object  "x - Back"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0-6 or x"
        if($result -eq "x")
        {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 7))
        {
            return [int]$result
        }
    }

    return $null
}

function Menu-MigrationWizConnectorTaskLoop([MigrationProxy.WebApi.MailboxConnector]$connector)
{
    $mailboxes = MWHelper-GetMailboxes -connector $connector

    # keep looping until specified to exit
    do
    {
        $action = Menu-MigrationWizConnectorTaskPrompt -connector $connector
        if($action -ne $null)
        {
            switch($action)
            {
                0 # Display mailbox migration status
                {
                    if($mailboxes -ne $null)
                    {
                        Action-MigrationWizConnectorTaskDisplayInfo -mailboxes $mailboxes
                    }
                    else
                    {
                        Write-Host
                        Write-Host -Object  "No mailboxes were found" -ForegroundColor Yellow
                    }
                }

                1 # Export statistics to CSV
                {
                    if($mailboxes -ne $null)
                    {
                        Action-MigrationWizConnectorTaskExportStats -mailboxes $mailboxes
                    }
                    else
                    {
                        Write-Host
                        Write-Host -Object  "No mailboxes were found" -ForegroundColor Yellow
                    }
                }

                2 # Submit mailboxes for migration
                {
                    if($mailboxes -ne $null)
                    {
                        Action-MigrationWizConnectorTaskSubmitMigration -mailboxes $mailboxes
                    }
                    else
                    {
                        Write-Host
                        Write-Host -Object  "No mailboxes were found" -ForegroundColor Yellow
                    }
                }

                3 # Stop all mailbox migrations
                {
                    if($mailboxes -ne $null)
                    {
                        Write-Host
                        if(Helper-PromptConfirmation -prompt "Are you sure you want to stop all migrations?")
                        {
                            Action-MigrationWizConnectorTaskStopMigrations -mailboxes $mailboxes
                        }
                    }
                    else
                    {
                        Write-Host
                        Write-Host -Object  "No mailboxes were found" -ForegroundColor Yellow
                    }
                }

                4 # Delete mailbox connector
                {
                    Write-Host
                    if(Helper-PromptConfirmation -prompt "Are you sure you want to delete this connector?")
                    {
                        Action-MigrationWizConnectorTaskDeleteConnector -connector $connector
                        return
                    }
                }

                5 # Configure mailbox connector
                {
                    Action-MigrationWizConnectorTaskConfigureConnector -connector $connector
                }

                6 # Reset configuration on all mailboxes to default
                {
                    Action-MigrationWizConnectorResetMailboxConfiguration -mailboxes $mailboxes
                }

                7 # Configure a public folder to shared mailbox migration
                {
                    Action-MigrationWizConnectorConfigurePublicFolderToSharedMailbox -connector $connector -mailboxes $mailboxes
                }
            }
        }
        else
        {
            return
        }
    }
    while($true)
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Mailbox task menu -> Configure a public folder to shared mailbox migraiton
######################################################################################################################################################

function Action-MigrationWizConnectorConfigurePublicFolderToSharedMailbox([MigrationProxy.WebApi.MailboxConnector]$connector, [MigrationProxy.WebApi.Mailbox[]]$mailboxes)
{
    Write-Host

    # Get the file containing the list of public folders
    $allPublicFoldersFileName = $null
    if ($script:allExchangeServerPublicFoldersFileName -ne $null)
    {
        $allPublicFoldersFileName = (Helper-PromptString -prompt "All public folders JSON file [$script:allExchangeServerPublicFoldersFileName]" -allowEmpty $true)
        if($allPublicFoldersFileName.Length -le 1)
        {
            $allPublicFoldersFileName = $script:allExchangeServerPublicFoldersFileName
        }
    }
    else
    {
        $allPublicFoldersFileName = (Helper-PromptString -prompt "All public folders JSON file" -allowEmpty $false)
    }

    # Ask for the domain in order to construct the email address for the shared mailbox
    $domain = (Helper-PromptString -prompt "Enter the domain for the shared mailbox email address (e.g. 'domain.com')" -allowEmpty $false)

    # Retrieve all objects from the mapping file
    Write-Host -Object "Converting mapping file to objects ..."
    $publicFoldersRaw = Get-Content $allPublicFoldersFileName | ConvertFrom-Json
    $publicFolders = @()

    # Convert from raw to PublicFolderData
    foreach($publicFolderRaw in $publicFoldersRaw)
    {
        # Create the public folder data object
        $publicFolderData = New-Object -TypeName BitTitan.ExchangeTools.PowerShell.Data.PublicFolderData

        # Set the public folder data
        $publicFolderData.FolderClass = $publicFolderRaw.FolderClass
        $publicFolderData.FolderSize = $publicFolderRaw.FolderSize
        $publicFolderData.Identity = $publicFolderRaw.Identity
        $publicFolderData.ParentPath = $publicFolderRaw.ParentPath
        $publicFolderData.IsValid = $publicFolderRaw.IsValid
        $publicFolderData.ItemCount = $publicFolderRaw.ItemCount
        $publicFolderData.MailEnabled = $publicFolderRaw.MailEnabled
        $publicFolderData.Name = $publicFolderRaw.Name
                

        # Add to the set of public folders
        $publicFolders += $publicFolderData
    }
    Write-Host -Object "Found $($publicFolders.Length) public folders in $allPublicFoldersFileName"

    $mailboxNamePrefix = (Helper-PromptString -prompt "Enter prefix for shared mailbox names [SharedMailbox-]" -allowEmpty $true)
    if ($mailboxNamePrefix -le 1)
    {
        $mailboxNamePrefix = "SharedMailbox-"
    }
    
    # Split the folders into 50GB shared mailboxes, grouping based on hierarchy
    Write-Host -Object "Splitting public folders into shared mailboxes"
    $publicFolderMailboxes = Split-PublicFoldersIntoMailboxes -MailboxSize 25 -PublicFolders $publicFolders -GroupingType SharedMailbox -MailboxNamePrefix $mailboxNamePrefix 

    # Output the public folder mailboxes
    $sharedMailboxesFileName = $env:TEMP + "SharedMailboxes.json"
    Write-Host -Object "Outputting shared mailboxes to $sharedMailboxesFileName"
    ConvertTo-Json -InputObject $publicFolderMailboxes -Compress -Depth 99 | Out-File -FilePath $sharedMailboxesFileName

    # Ask the user whether we should attempt to create the Shared Mailboxes in Office 365
    $createO365Mailboxes = (Helper-PromptConfirmation -prompt "Attempt to create $($publicFolderMailboxes.Length) shared mailboxes in Office 365?")

    # Create the mailboxes in Office 365
    if ($createO365Mailboxes)
    {
        try
        {
            # Create the O365 PS Session
            Office365Helper-ConnectRemotePowerShell
            $credentials = Office365Helper-GetCredentials

            # Attempt to create each Shared Mailbox
            Write-Host -Object "Creating $($publicFolderMailboxes.Length) shared mailboxes in Office 365"
            foreach($publicFolderMailbox in $publicFolderMailboxes)
            {
                # Set the alias
                $alias = $publicFolderMailbox.MailboxName

                # Look for an existing mailbox
                $existingMailbox = Get-Mailbox -Identity $alias -ErrorAction SilentlyContinue
                if ($existingMailbox -ne $null)
                {
                    if ($existingMailbox.IsShared)
                    {
                        # Use the existing shared mailbox if one already exists
                        Write-Warning -Message "  Shared Mailbox already exists for alias $alias. Skipping creation."
                        continue
                    }

                    # If there is already a non-shared mailbox for the given alias, then throw an exception
                    throw "Error: Non-Shared Mailbox already exists for alias $alias."
                }

                # Create a new Shared Mailbox for the given alias
                Write-Host -Object "  Attempting to created shared mailbox with alias $alias in Office 365 Exchange Online"
                $o365Mailbox = New-Mailbox -Name "$alias" -DisplayName "$alias" -Alias "$alias" -Shared
                $o365Mailbox | Add-MailboxPermission -User $credentials.UserName -AccessRights FullAccess -InheritanceType All
                Write-Host -Object "  Created shared mailbox $($o365Mailbox.PrimarySmtpAddress) in Office 365 Exchange Online"
            }
        }
        finally
        {
            Get-PSSession | Remove-PSSession
        }
    }

    # Create mailboxes for each chunk of folders
    Write-Host -Object "Creating $($publicFolderMailboxes.Length) project items in MigrationWiz"
    foreach($publicFolderMailbox in $publicFolderMailboxes)
    {
        $mailbox = New-Object -TypeName MigrationProxy.WebApi.Mailbox
        $mailbox.ConnectorId = $connector.Id
        $mailbox.ExportEmailAddress = ""
        $mailbox.ExportPassword = ""
        $mailbox.ExportUserName = ""
        $mailbox.ImportEmailAddress = "$($publicFolderMailbox.MailboxName)@$domain"
        $mailbox.ImportPassword = ""
        $mailbox.ImportUserName = ""
        $mailbox.PublicFolderPath = $publicFolderMailbox.RootIdentity
        $mailbox.FolderFilter = $publicFolderMailbox.FolderFilterString

        # Add the mailbox to the connector
        MW-AddMailbox -mailbox $mailbox
        Write-Host -Object "Created project item for shared mailbox $($mailbox.ImportEmailAddress)"

        # TODO: attempt to create the shared mailbox if it doesn't exist in O365

        # Add the newly created mailbox to the list of mailboxes
        $mailboxes += $mailbox
    }

    # Return all of the mailboxes
    return $mailboxes
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Mailbox task menu -> Display mailbox migration status
######################################################################################################################################################

function Action-MigrationWizConnectorTaskDisplayInfo([MigrationProxy.WebApi.Mailbox[]]$mailboxes)
{
    $count = 0
    $totalCount = 0
    $totalSpeed = 0
    $totalItems = 0

    Write-Host
    foreach($mailbox in $mailboxes)
    {
        $count++
        Write-Progress -Activity ("Retrieving mailbox information for " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)

        Action-MigrationWizConnectorTaskDisplayInfoStatus -mailbox $mailbox
        $statsInfo = Action-MigrationWizConnectorTaskDisplayInfoSpeeds -mailbox $mailbox

        if($statsInfo[0] -gt 0 -and $statsInfo[1] -gt 0)
        {
            $totalCount++
            $totalSpeed += $statsInfo[0]
            $totalItems += $statsInfo[1]
        }
    }

    Write-Host
    if($totalCount -gt 0)
    {
        $totalSpeed = $totalSpeed / $totalCount
        $totalItems = $totalItems / $totalCount

        Write-Host -Object  "Average speed was $totalSpeed MB/hour and $totalItems items/hour"
    }
    else
    {
        Write-Host -Object  "Migrations did not run long enough to compile statistics" -ForegroundColor Yellow
    }
}

function Action-MigrationWizConnectorTaskDisplayInfoSpeeds([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $returnSpeed = 0
    $returnCount = 0

    $stats = MWHelper-GetMailboxStatistics -mailbox $mailbox
    if($stats[0])
    {
        $folderSuccessSize = $stats[1]
        $calendarSuccessSize = $stats[2]
        $contactSuccessSize = $stats[3]
        $mailSuccessSize = $stats[4]
        $taskSuccessSize = $stats[5]
        $noteSuccessSize = $stats[6]
        $journalSuccessSize = $stats[7]
        $totalSuccessSize = $stats[8]

        $folderSuccessCount = $stats[9]
        $calendarSuccessCount = $stats[10]
        $contactSuccessCount = $stats[11]
        $mailSuccessCount = $stats[12]
        $taskSuccessCount = $stats[13]
        $noteSuccessCount = $stats[14]
        $journalSuccessCount = $stats[15]
        $totalSuccessCount = $stats[16]

        $folderErrorSize = $stats[17]
        $calendarErrorSize = $stats[18]
        $contactErrorSize = $stats[19]
        $mailErrorSize = $stats[20]
        $taskErrorSize = $stats[21]
        $noteErrorSize = $stats[22]
        $journalErrorSize = $stats[23]
        $totalErrorSize = $stats[24]

        $folderErrorCount = $stats[25]
        $calendarErrorCount = $stats[26]
        $contactErrorCount = $stats[27]
        $mailErrorCount = $stats[28]
        $taskErrorCount = $stats[29]
        $noteErrorCount = $stats[30]
        $journalErrorCount = $stats[31]
        $totalErrorCount = $stats[32]

        $totalExportActiveDuration = $stats[33]
        $totalExportPassiveDuration = $stats[34]
        $totalImportActiveDuration = $stats[35]
        $totalImportPassiveDuration = $stats[36]

        $totalExportSpeed = $stats[37]
        $totalExportCount = $stats[38]

        $totalImportSpeed = $stats[39]
        $totalImportCount = $stats[40]

        if($totalExportSpeed -lt $totalImportSpeed)
        {
            Write-Host -Object  ("  " + $totalExportSpeed.ToString($precision) + " MB/hour (" + $totalExportCount.ToString($precision) + " items/hour) export in " + $totalExportActiveDuration.ToString($precision) + " minutes") -NoNewLine
            if($totalExportActiveDuration -gt 60)
            {
                $returnSpeed = $totalExportSpeed;
                $returnCount = $totalExportCount;
            }
            else
            {
                Write-Host -Object  " (Not accurate)" -NoNewLine -ForegroundColor Yellow
            }
        }
        else
        {
            Write-Host -Object  ("  " + $totalImportSpeed.ToString($precision) + " MB/hour (" + $totalImportCount.ToString($precision) + " items/hour) import in " + $totalImportActiveDuration.ToString($precision) + " minutes") -NoNewLine
            if($totalImportActiveDuration -gt 60)
            {
                $returnSpeed = $totalImportSpeed;
                $returnCount = $totalImportCount;
            }
            else
            {
                Write-Host -Object  " (Not accurate)" -NoNewLine -ForegroundColor Yellow
            }
        }

        Write-Host
    }
    else
    {
        Write-Host -Object  "  No migration statistics found" -ForegroundColor Yellow
    }

    return @($returnSpeed, $returnCount)
}

function Action-MigrationWizConnectorTaskDisplayInfoStatus([MigrationProxy.WebApi.Mailbox]$mailbox)
{
    $status = "NotMigrated"
    $color = [System.ConsoleColor]::Gray

    $migration = MW-GetMailboxMigrationLatest -mailbox $mailbox
    if($migration -ne $null)
    {
        $status = $migration.Status
    }

    switch($status)
    {
        "NotMigrated"
        {
            $color = [System.ConsoleColor]::Yellow
        }

        "Submitted"
        {
            $color = [System.ConsoleColor]::White
        }

        "WaitingForEndUser"
        {
            $color = [System.ConsoleColor]::Yellow
        }

        "Queued"
        {
            $color = [System.ConsoleColor]::White
        }

        "Processing"
        {
            $color = [System.ConsoleColor]::White
        }

        "Completed"
        {
        }

        "Failed"
        {
            $color = [System.ConsoleColor]::Red
        }

        "Stopping"
        {
            $color = [System.ConsoleColor]::Yellow
        }

        "Stopped"
        {
            $color = [System.ConsoleColor]::Red
        }

        "MaximumTransferReached"
        {
            $color = [System.ConsoleColor]::Yellow
        }
    }

    Write-Host -Object  ($mailbox.ExportEmailAddress + " -> " + $mailbox.ImportEmailAddress) -NoNewLine
    Write-Host -Object  (" (" + $status + ")") -ForegroundColor $color
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Mailbox task menu -> Export statistics to CSV
######################################################################################################################################################

function Action-MigrationWizConnectorTaskExportStats([MigrationProxy.WebApi.Mailbox[]]$mailboxes)
{
    $statsFilename = Helper-GenerateRandomTempFilename -identifier "Statistics"
    $errorsFilename = Helper-GenerateRandomTempFilename -identifier "Errors"

    Write-Host
    Write-Host -Object  ("Exporting connector statistics to " + $statsFilename)
    Write-Host -Object  ("Exporting connector errors to " + $errorsFilename)

    $statsLine = "Mailbox Id,Source Email Address,Destination Email Address"
    $statsLine += ",Folders Success Count,Folders Success Size (bytes),Folders Error Count,Folders Error Size (bytes)"
    $statsLine += ",Calendars Success Count,Calendars Success Size (bytes),Calendars Error Count,Calendars Error Size (bytes)"
    $statsLine += ",Contacts Success Count,Contacts Success Size (bytes),Contacts Error Count,Contacts Error Size (bytes)"
    $statsLine += ",Email Success Count,Email Success Size (bytes),Email Error Count,Email Error Size (bytes)"
    $statsLine += ",Tasks Success Count,Tasks Success Size (bytes),Tasks Error Count,Tasks Error Size (bytes)"
    $statsLine += ",Notes Success Count,Notes Success Size (bytes),Notes Error Count,Notes Error Size (bytes)"
    $statsLine += ",Journals Success Count,Journals Success Size (bytes),Journals Error Count,Journals Error Size (bytes)"
    $statsLine += ",Total Success Count,Total Success Size (bytes),Total Error Count,Total Error Size (bytes)"
    $statsLine += ",Source Active Duration (minutes),Source Passive Duration (minutes),Source Data Speed (MB/hour),Source Item Speed (items/hour)"
    $statsLine += ",Destination Active Duration (minutes),Destination Passive Duration (minutes),Destination Data Speed (MB/hour),Destination Item Speed (items/hour)"
    $statsLine += ",Migrations Performed,Last Migration Type,Last Status,Last Status Details"
    $statsLine += "`r`n"

    $errorsLine = "Mailbox Id,Source Email Address,Destination Email Address,Type,Date,Size (bytes),Error,Subject`r`n"

    $file = New-Item -Path $statsFilename -ItemType file -force -value $statsLine
    $file = New-Item -Path $errorsFilename -ItemType file -force -value $errorsLine

    $count = 0

    foreach($mailbox in $mailboxes)
    {
        $count++

        Write-Progress -Activity ("Retrieving mailbox information for " + $connector.Name + " (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)
        $stats = MWHelper-GetMailboxStatistics -mailbox $mailbox
        $migrations = MW-GetMailboxHistory -mailbox $mailbox
        $errors = MW-GetMailboxErrors -mailbox $mailbox

        $statsLine = $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress

        $folderSuccessSize = $stats[1]
        $calendarSuccessSize = $stats[2]
        $contactSuccessSize = $stats[3]
        $mailSuccessSize = $stats[4]
        $taskSuccessSize = $stats[5]
        $noteSuccessSize = $stats[6]
        $journalSuccessSize = $stats[7]
        $totalSuccessSize = $stats[8]

        $folderSuccessCount = $stats[9]
        $calendarSuccessCount = $stats[10]
        $contactSuccessCount = $stats[11]
        $mailSuccessCount = $stats[12]
        $taskSuccessCount = $stats[13]
        $noteSuccessCount = $stats[14]
        $journalSuccessCount = $stats[15]
        $totalSuccessCount = $stats[16]

        $folderErrorSize = $stats[17]
        $calendarErrorSize = $stats[18]
        $contactErrorSize = $stats[19]
        $mailErrorSize = $stats[20]
        $taskErrorSize = $stats[21]
        $noteErrorSize = $stats[22]
        $journalErrorSize = $stats[23]
        $totalErrorSize = $stats[24]

        $folderErrorCount = $stats[25]
        $calendarErrorCount = $stats[26]
        $contactErrorCount = $stats[27]
        $mailErrorCount = $stats[28]
        $taskErrorCount = $stats[29]
        $noteErrorCount = $stats[30]
        $journalErrorCount = $stats[31]
        $totalErrorCount = $stats[32]

        $totalExportActiveDuration = $stats[33]
        $totalExportPassiveDuration = $stats[34]
        $totalImportActiveDuration = $stats[35]
        $totalImportPassiveDuration = $stats[36]

        $totalExportSpeed = $stats[37]
        $totalExportCount = $stats[38]

        $totalImportSpeed = $stats[39]
        $totalImportCount = $stats[40]

        $statsLine += "," + $folderSuccessCount + "," + $folderSuccessSize + "," + $folderErrorCount + "," + $folderErrorSize
        $statsLine += "," + $calendarSuccessCount + "," + $calendarSuccessSize + "," + $calendarErrorCount + "," + $calendarErrorSize
        $statsLine += "," + $contactSuccessCount + "," + $contactSuccessSize + "," + $contactErrorCount + "," + $contactErrorSize
        $statsLine += "," + $mailSuccessCount + "," + $mailSuccessSize + "," + $mailErrorCount + "," + $mailErrorSize
        $statsLine += "," + $taskSuccessCount + "," + $taskSuccessSize + "," + $taskErrorCount + "," + $taskErrorSize
        $statsLine += "," + $noteSuccessCount + "," + $noteSuccessSize + "," + $noteErrorCount + "," + $noteErrorSize
        $statsLine += "," + $journalSuccessCount + "," + $journalSuccessSize + "," + $journalErrorCount + "," + $journalErrorSize
        $statsLine += "," + $totalSuccessCount + "," + $totalSuccessSize + "," + $totalErrorCount + "," + $totalErrorSize
        $statsLine += "," + $totalExportActiveDuration + "," + $totalExportPassiveDuration + "," + $totalExportSpeed + "," + $totalExportCount
        $statsLine += "," + $totalImportActiveDuration + "," + $totalImportPassiveDuration + "," + $totalImportSpeed + "," + $totalImportCount

        if($migrations -ne $null)
        {
            $latest = $migrations[$migrations.Length-1]
            $statsLine += "," + $migrations.Length + "," + $latest.Type + "/" + $latest.LicenseSku + "," + $latest.Status

            if($latest.FailureMessage -ne $null)
            {
                $statsLine +=  ',"' + $latest.FailureMessage.Replace('"', "'") + '"'
            }
            else
            {
                $statsLine +=  ","
            }
        }
        else
        {
            $statsLine += ",,,NotMigrated,"
        }

        if($errors -ne $null)
        {
            if($errors.Length -ge 1)
            {
                foreach($error in $errors)
                {
                    $errorsLine = $mailbox.Id.ToString() + "," + $mailbox.ExportEmailAddress + "," + $mailbox.ImportEmailAddress
                    $errorsLine += "," + $error.Type.ToString()
                    $errorsLine += "," + $error.CreateDate.ToString("M/d/yyyy h:mm tt")
                    $errorsLine += "," + $error.ItemSize

                    if($error.Message -ne $null)
                    {
                        $errorsLine +=  ',"' + $error.Message.Replace('"', "'") + '"'
                    }
                    else
                    {
                        $errorsLine +=  ","
                    }

                    if($error.ItemSubject -ne $null)
                    {
                        $errorsLine +=  ',"' + $error.ItemSubject.Replace('"', "'") + '"'
                    }
                    else
                    {
                        $errorsLine +=  ","
                    }
                    Add-Content -Path $errorsFilename -Value $errorsLine
                }
            }
        }

        Add-Content -Path $statsFilename -Value $statsLine
    }
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Mailbox task menu -> Delete mailbox connector
######################################################################################################################################################

function Action-MigrationWizConnectorTaskDeleteConnector([MigrationProxy.WebApi.MailboxConnector]$connector)
{
    Write-Host
    Write-Host -Object  "Deleting mailbox connector $($connector.Name) ..."

    $continue = (Helper-PromptConfirmation -prompt "Are you sure you wish to continue?")
    if($continue)
    {
        MW-RemoveMailboxConnector -connector $connector
    }
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Mailbox task menu -> Configure mailbox connector
######################################################################################################################################################

function Action-MigrationWizConnectorTaskConfigureConnector([MigrationProxy.WebApi.MailboxConnector]$connector)
{
    $configType = Menu-MigrationWizConfigureConnectorPrompt
    if($configType -ne $null)
    {
        Write-Progress -Activity ("Updating mailbox connector " + $connector.Name) -Status " "

        switch($configType)
        {
            0 # Reset configuration to default
            {
                $connector.MaximumSimultaneousMigrations = 10
                $connector.ItemStartDate = Get-Date -Year 1901 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                $connector.ItemEndDate = Get-Date -Year 9999 -Month 12 -Day 31 -Hour 23 -Minute 59 -Second 59
                $connector.MaximumItemFailures = 100
                $connector.PurgePeriod = 10
                $connector.FolderFilter = ""
                $connector.AdvancedOptions = ""

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Calendar)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Calendar
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Contact)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Contact
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Mail)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Mail
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Journal)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Journal
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Note)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Note
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Task)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Task
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Rule)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Rule
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyAdminComplete)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyAdminComplete
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyAdminFailed)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyAdminFailed
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyExportComplete)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyExportComplete
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyExportFailed)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyExportFailed
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyImportComplete)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyImportComplete
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyImportFailed)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyImportFailed
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::DoNotContinueFromLastKnownState)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotContinueFromLastKnownState
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::DoNotQuarantineFatalErrors)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotQuarantineFatalErrors
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::DoNotSearchImportForDuplicates)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotSearchImportForDuplicates
                }

                if($connector.Flags -band [MigrationProxy.WebApi.MailboxFlags]::DoNotRetryErrors)
                {
                    $connector.Flags = $connector.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotRetryErrors
                }
            }

            1 # Migrate email only excluding the inbox
            {
                $connector.ItemStartDate = Get-Date -Year 1901 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                $connector.ItemEndDate = Get-Date -Year 9999 -Month 12 -Day 31 -Hour 23 -Minute 59 -Second 59
                $connector.FolderFilter = "^Inbox$"

                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Calendar
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Contact
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Journal
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Note
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Task
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Rule

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Mail)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Mail
                }
            }

            2 # Migrate items older than 2 months within inbox only
            {
                $connector.ItemStartDate = Get-Date -Year 1901 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                $connector.ItemEndDate = (Get-Date).AddMonths(-2)
                $connector.FolderFilter = "^(?!Inbox$)"

                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Calendar
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Contact
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Journal
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Note
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Task
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Rule

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Mail)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Mail
                }
            }

            8 # Migrate email only
            {
                $connector.ItemStartDate = Get-Date -Year 1901 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                $connector.ItemEndDate = Get-Date -Year 9999 -Month 12 -Day 31 -Hour 23 -Minute 59 -Second 59
                $connector.FolderFilter = ""

                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Calendar
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Contact
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Journal
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Note
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Task
                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Rule

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Mail)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Mail
                }
            }

            9 # Migrate non-email items (calendar, contacts, tasks, journals, notes, rules)
            {
                $connector.ItemStartDate = Get-Date -Year 1901 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                $connector.ItemEndDate = Get-Date -Year 9999 -Month 12 -Day 31 -Hour 23 -Minute 59 -Second 59
                $connector.FolderFilter = ""

                $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bor [MigrationProxy.WebApi.MailboxItemTypes]::Mail

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Calendar)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Calendar
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Contact)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Contact
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Journal)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Journal
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Note)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Note
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Task)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Task
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Rule)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Rule
                }
            }

            10 # Migrate all items
            {
                $connector.ItemStartDate = Get-Date -Year 1901 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
                $connector.ItemEndDate = Get-Date -Year 9999 -Month 12 -Day 31 -Hour 23 -Minute 59 -Second 59
                $connector.FolderFilter = ""

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Mail)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Mail
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Calendar)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Calendar
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Contact)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Contact
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Journal)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Journal
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Note)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Note
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Task)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Task
                }

                if($connector.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Rule)
                {
                    $connector.DisabledMailboxItemTypes = $connector.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Rule
                }
            }
        }

        MW-SetMailboxConnector -connector $connector
    }
}

function Menu-MigrationWizConfigureConnectorPrompt()
{
    Write-Host
    Write-Host -Object  ("How would you like to configure the mailbox connector options:") -ForegroundColor Yellow
    Write-Host
    Write-Host -Object  "0  - Reset configuration to default" #kill
    Write-Host -Object  "1  - Migrate email only excluding the inbox" #kill
    Write-Host -Object  "2  - Migrate email older than 1 month"
    Write-Host -Object  "3  - Migrate email older than 3 months"
    Write-Host -Object  "8  - Migrate email only" #kill
    Write-Host -Object  "9  - Migrate non-email items (calendar, contacts, tasks, journals, notes, rules)" #kill
    Write-Host -Object  "10 - Migrate all items"
    Write-Host -Object  "x  - Back"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0-10 or x"
        if($result -eq "x")
        {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 10))
        {
            return [int]$result
        }
    }

    return $null
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Mailbox task menu -> Submit mailboxes for migration
######################################################################################################################################################

function Action-MigrationWizConnectorTaskSubmitMigration([MigrationProxy.WebApi.Mailbox[]]$mailboxes)
{
    $count = 0
    $mailboxesToSubmit = $null

    $statusAction = Menu-MigrationWizGetMailboxStatusPrompt
    if($statusAction -ne $null)
    {
        $licenseAction = Menu-MigrationWizGetLicensePrompt
        if($licenseAction -ne $null)
        {
            $licenseType = $licenseAction[0]

            Write-Host
            if($statusAction -eq 4)
            {
                do
                {
                    $emailAddress = (Helper-PromptString -prompt "Email address of mailbox to submit (Press enter when done)" -allowEmpty $true)
                    if($emailAddress.Length -ge 1)
                    {
                        $mailboxesToSubmit += @($emailAddress)
                    }
                }
                while($emailAddress.Length -ge 1)

                if($mailboxesToSubmit -eq $null -or $mailboxesToSubmit.Length -le 0)
                {
                    return
                }
            }

            Write-Host
            Write-Host -Object "Submitting mailboxes for migration ..."

            foreach($mailbox in $mailboxes)
            {
                $submit = $false
                $status = "NotMigrated"

                $count++
                Write-Progress -Activity ("Submitting mailboxes for migration (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)

                if($statusAction -ne 4)
                {
                    $migration = MW-GetMailboxMigrationLatest -mailbox $mailbox
                    if($migration -ne $null)
                    {
                        $status = $migration.Status
                    }
                }

                switch($status)
                {
                    "NotMigrated"
                    {
                        if($statusAction -eq 0 -or $statusAction -eq 1 -or $statusAction -eq 5)
                        {
                            $submit = $true
                        }
                        elseif($statusAction -eq 4 -and $mailboxesToSubmit -ne $null -and $mailboxesToSubmit.Length -ge 1)
                        {
                            if((Helper-StringInArray -toFind $mailbox.ExportEmailAddress -stringArray $mailboxesToSubmit) -or (Helper-StringInArray -toFind $mailbox.ImportEmailAddress -stringArray $mailboxesToSubmit))
                            {
                                $submit = $true
                            }
                        }
                    }

                    "Completed"
                    {
                        if($statusAction -eq 0)
                        {
                            $submit = $true
                        }
                        elseif($statusAction -eq 3)
                        {
                            $stats = MWHelper-GetMailboxStatistics -mailbox $mailbox
                            if($stats[32] -ge 1)
                            {
                                $submit = $true
                            }
                        }
                        elseif($statusAction -eq 4 -and $mailboxesToSubmit -ne $null -and $mailboxesToSubmit.Length -ge 1)
                        {
                            if((Helper-StringInArray -toFind $mailbox.ExportEmailAddress -stringArray $mailboxesToSubmit) -or (Helper-StringInArray -toFind $mailbox.ImportEmailAddress -stringArray $mailboxesToSubmit))
                            {
                                $submit = $true
                            }
                        }
                    }

                    "Failed"
                    {
                        if($statusAction -eq 0 -or $statusAction -eq 2 -or $statusAction -eq 5)
                        {
                            $submit = $true
                        }
                        elseif($statusAction -eq 4 -and $mailboxesToSubmit -ne $null -and $mailboxesToSubmit.Length -ge 1)
                        {
                            if((Helper-StringInArray -toFind $mailbox.ExportEmailAddress -stringArray $mailboxesToSubmit) -or (Helper-StringInArray -toFind $mailbox.ImportEmailAddress -stringArray $mailboxesToSubmit))
                            {
                                $submit = $true
                            }
                        }
                    }

                    "Stopped"
                    {
                        if($statusAction -eq 0 -or $statusAction -eq 5)
                        {
                            $submit = $true
                        }
                        elseif($statusAction -eq 4 -and $mailboxesToSubmit -ne $null -and $mailboxesToSubmit.Length -ge 1)
                        {
                            if((Helper-StringInArray -toFind $mailbox.ExportEmailAddress -stringArray $mailboxesToSubmit) -or (Helper-StringInArray -toFind $mailbox.ImportEmailAddress -stringArray $mailboxesToSubmit))
                            {
                                $submit = $true
                            }
                        }
                    }

                    "MaximumTransferReached"
                    {
                        if($statusAction -eq 0 -or $statusAction -eq 5)
                        {
                            $submit = $true
                        }
                        elseif($statusAction -eq 4 -and $mailboxesToSubmit -ne $null -and $mailboxesToSubmit.Length -ge 1)
                        {
                            if((Helper-StringInArray -toFind $mailbox.ExportEmailAddress -stringArray $mailboxesToSubmit) -or (Helper-StringInArray -toFind $mailbox.ImportEmailAddress -stringArray $mailboxesToSubmit))
                            {
                                $submit = $true
                            }
                        }
                    }
                }

                if($submit)
                {
                    MW-SubmitMigration -mailbox $mailbox -licenseType $licenseType
                }
            }
        }
    }
}

function Menu-MigrationWizGetMailboxStatusPrompt()
{
    Write-Host
    Write-Host -Object ("Which mailboxes would you like to submit:") -ForegroundColor Yellow
    Write-Host
    Write-Host -Object "0 - All mailboxes"
    Write-Host -Object "1 - Not migrated mailboxes"
    Write-Host -Object "2 - Failed mailboxes"
    Write-Host -Object "3 - Successful mailboxes that contain errors"
    Write-Host -Object "4 - Specify the email address"
    Write-Host -Object "5 - All mailboxes that were not successful"
    Write-Host -Object "x - Back"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0-5 or x"
        if($result -eq "x")
        {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 5))
        {
            return [int]$result
        }
    }

    return $null
}

function Menu-MigrationWizGetLicensePrompt()
{
    Write-Host
    Write-Host -Object ("What type of migration would you like to perform:") -ForegroundColor Yellow
    Write-Host
    Write-Host -Object "0 - Mailbox migration (including delta pass if previously migrated)"
    Write-Host -Object "1 - Verify credentials"
    Write-Host -Object "2 - Retry mailbox migration errors"
    Write-Host -Object "3 - Trial mailbox migration"
    Write-Host -Object "x - Back"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0-4 or x"
        if($result -eq "x")
        {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 4))
        {
            switch([int]$result)
            {
                0
                {
                    $licenseType = [MigrationProxy.WebApi.MailboxQueueTypes]::Full
                }

                1
                {
                    $licenseType = [MigrationProxy.WebApi.MailboxQueueTypes]::Verification
                }

                2
                {
                    $licenseType = [MigrationProxy.WebApi.MailboxQueueTypes]::Repair
                }

                3
                {
                    $licenseType = [MigrationProxy.WebApi.MailboxQueueTypes]::Trial
                }
            }

            return @($licenseType)
        }
    }

    return $null
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Mailbox task menu -> Stop all mailboxes migrations
######################################################################################################################################################

function Action-MigrationWizConnectorTaskStopMigrations([MigrationProxy.WebApi.Mailbox[]]$mailboxes)
{
    Write-Host
    Write-Host -Object "Stopping all mailbox migrations ..."

    $runningLast = 0
    $stoppingLast = 0
    $notRunningLast = 0

    while($true)
    {
        $runningCount = 0
        $stoppingCount = 0
        $notRunningCount = 0

        foreach($mailbox in $mailboxes)
        {
            Write-Progress -Activity ("Checking mailbox status ... " + $runningLast + " running / " + $stoppingLast + " stopping / " + $notRunningLast + " not running") -Status $mailbox.ExportEmailAddress

            $migration = MW-GetMailboxMigrationLatest -mailbox $mailbox
            if($migration -ne $null)
            {
                switch($migration.Status)
                {
                    Submitted
                    {
                        $runningCount++
                    }

                    WaitingForEndUser
                    {
                        $runningCount++
                        MW-StopMigration -migration $migration
                    }

                    Queued
                    {
                        $runningCount++
                    }

                    Processing
                    {
                        $runningCount++
                        MW-StopMigration -migration $migration
                    }

                    Stopping
                    {
                        $stoppingCount++
                    }

                    Completed
                    {
                        $notRunningCount++
                    }

                    Failed
                    {
                        $notRunningCount++
                    }

                    MaximumTransferReached
                    {
                        $notRunningCount++
                    }

                    Stopped
                    {
                        $notRunningCount++
                    }
                }
            }
            else
            {
                $notRunningCount++
            }
        }

        $runningLast = $runningCount
        $stoppingLast = $stoppingCount
        $notRunningLast = $notRunningCount

        if($runningCount -le 0 -and $stoppingCount -le 0)
        {
            return
        }
        else
        {
            Write-Progress -Activity ("Checking mailbox status ... " + $runningLast + " running / " + $stoppingLast + " stopping / " + $notRunningLast + " not running") -Status ("Will retry in " + $stopSleepSeconds + " seconds")
            Start-Sleep -Seconds $stopSleepSeconds
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage mailbox connectors -> Connector -> Reset configuration on all mailboxes to default
######################################################################################################################################################

function Action-MigrationWizConnectorResetMailboxConfiguration([MigrationProxy.WebApi.Mailbox[]]$mailboxes)
{
    $count = 0

    Write-Host
    Write-Host -Object "Resetting mailbox configurations ..."

    foreach($mailbox in $mailboxes)
    {
        $count++
        Write-Progress -Activity ("Resetting mailbox configuration (" + $count + "/" + $mailboxes.Length + ")") -Status $mailbox.ExportEmailAddress -PercentComplete ($count/$mailboxes.Length*100)

        $mailbox.FolderFilter = ""
        $mailbox.AdvancedOptions = ""

        if($mailbox.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Calendar)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotMigrateCalendar
            $mailbox.DisabledMailboxItemTypes = $mailbox.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Calendar
        }

        if($mailbox.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Contact)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotMigrateContact
            $mailbox.DisabledMailboxItemTypes = $mailbox.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Contact
        }

        if($mailbox.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Mail)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotMigrateMail
            $mailbox.DisabledMailboxItemTypes = $mailbox.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Mail
        }

        if($mailbox.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Journal)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotMigrateJournal
            $mailbox.DisabledMailboxItemTypes = $mailbox.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Journal
        }

        if($mailbox.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Note)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotMigrateNote
            $mailbox.DisabledMailboxItemTypes = $mailbox.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Note
        }

        if($mailbox.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Task)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotMigrateTask
            $mailbox.DisabledMailboxItemTypes = $mailbox.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Task
        }

        if($mailbox.DisabledMailboxItemTypes -band [MigrationProxy.WebApi.MailboxItemTypes]::Rule)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotMigrateRule
            $mailbox.DisabledMailboxItemTypes = $mailbox.DisabledMailboxItemTypes -bxor [MigrationProxy.WebApi.MailboxItemTypes]::Rule
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyAdminComplete)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyAdminComplete
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyAdminFailed)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyAdminFailed
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyExportComplete)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyExportComplete
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyExportFailed)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyExportFailed
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyImportComplete)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyImportComplete
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::NotifyImportFailed)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::NotifyImportFailed
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::DoNotContinueFromLastKnownState)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotContinueFromLastKnownState
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::DoNotQuarantineFatalErrors)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotQuarantineFatalErrors
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::DoNotSearchImportForDuplicates)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotSearchImportForDuplicates
        }

        if($mailbox.Flags -band [MigrationProxy.WebApi.MailboxFlags]::DoNotRetryErrors)
        {
            $mailbox.Flags = $mailbox.Flags -bxor [MigrationProxy.WebApi.MailboxFlags]::DoNotRetryErrors
        }

        MW-SetMailbox -mailbox $mailbox
    }
}

######################################################################################################################################################
# Main menu -> Manage G Suite
######################################################################################################################################################

function Menu-GooglePrompt()
{
    Write-Host
    Write-Host -Object ("Select a G Suite task to perform:") -ForegroundColor Yellow
    Write-Host
    Write-Host -Object "0 - Create MigrationWiz mailbox import file"
    Write-Host -Object "1 - Set email forwarding for mailbox (Only works if the address has been verified to be a forwarding address in the account or it is in the account's primary domain or a subdomain)"
    Write-Host -Object "2 - Remove email forwarding for mailbox"
    Write-Host -Object "3 - Export users to CSV"
    Write-Host -Object "4 - Revoke access token"
    Write-Host -Object "x - Back"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0-4 or x"
        if($result -eq "x")
        {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 4))
        {
            return [int]$result
        }
    }

    return $null
}

function Menu-GoogleLoop()
{
    # keep looping until specified to exit
    do
    {
        $action = Menu-GooglePrompt
        if($action -ne $null)
        {
            switch($action)
            {
                0 # Create MigrationWiz mailbox import file
                {
                    Action-GoogleExportMailboxes
                }

                1 # Set email forwarding for mailbox
                {
                    Action-GoogleSetMailboxForward
                }

                2 # Remove email forwarding for mailbox
                {
                    Action-GoogleRemoveMailboxForward
                }

                3 # Export users to CSV
                {
                    Action-GoogleExportUsers
                }

                4 # Revoke access token
                {
                    Action-GoogleRevokeOAuth2AccessToken
                }
            }
        }
        else
        {
            return
        }
    }
    while($true)
}

######################################################################################################################################################
# Main menu -> Manage G Suite -> Create MigrationWiz mailbox import file
######################################################################################################################################################

function Action-GoogleExportMailboxes()
{
    Write-Host

    do
    {
        $domainName = Read-Host -Prompt "Enter the G Suite domain to export"
    }
    while($domainName.Length -le 0)

    $users = GoogleHelper-GetDomainUsers -domainName $domainName
    if($users -ne $null)
    {
        $count = 0
        $csv = "Source Email,Source UserName,Source Password,Destination Email,Destination UserName,Destination Password`r`n"
        $filename = Helper-GenerateRandomTempFilename -identifier "GoogleMailboxes"
        $file = New-Item -Path $filename -ItemType file -force -value $csv

        Write-Host -Object ("Creating MigrationWiz import file " + $filename)

        foreach($user in $users.Entries)
        {
            $emailAddress = $user.Login.UserName + "@" + $domainName

            $count++
            Write-Progress -Activity ("Exporting G Suite mailboxes for " + $domainName + " (" + $count + ")") -Status $emailAddress

            $csv = ""
            $csv += '"' + $emailAddress + '"' + ','		# Source Email
            $csv += '"",'								# Source UserName
            $csv += '"",'								# Source Password
            $csv += '"' + $emailAddress + '"' + ','		# Destination Email
            $csv += '"",'								# Destination UserName
            $csv += '""'						 		# Destination Password

            Add-Content -Path $filename -Value $csv
        }
    }
    else
    {
        Write-Host
        Write-Host -Object "No users found within the domain specified.  Possible causes include:" -ForegroundColor Red
        Write-Host -Object "   1) no users" -ForegroundColor Red
        Write-Host -Object "   2) invalid domain specified" -ForegroundColor Red
        Write-Host -Object "   3) invalid credentials specified" -ForegroundColor Red
    }
}

######################################################################################################################################################
# Main menu -> Manage G Suite -> Set email forwarding for mailbox
######################################################################################################################################################

function Action-GoogleSetMailboxForward()
{
    $count = 0

    Write-Host

    $importConfirm = (Helper-PromptConfirmation -prompt "Would you like to import a list from a file?")
    if($importConfirm)
    {
        $importFilename = (Helper-PromptString -prompt "Enter the full path to import file (Press enter to create one)" -allowEmpty $true)
        if($importFilename -eq "")
        {
            # create new import file
            $importFilename = Helper-GenerateRandomTempFilename -identifier "GoogleForwardingImport"
            $csv = "Email Address,Forwarding Address`r`n"
            $file = New-Item -Path $importFilename -ItemType file -force -value $csv

            # open file for editing
            Start-Process -FilePath $importFilename

            do
            {
                $importConfirm = (Helper-PromptConfirmation -prompt "Are you done editing the import file?")
            }
            while(-not $importConfirm)
        }

        # read csv file
        $users = Import-Csv -Path $importFilename
        foreach($user in $users)
        {
            $count++

            $emailAddress = $user.'Email Address'
            $targetAddress = $user.'Forwarding Address'

            if($emailAddress -ne $null -and $emailAddress -ne "" -and $targetAddress -ne $null -and $targetAddress -ne "")
            {
                Write-Progress -Activity ("Setting mailbox forward (" + $count + ")") -Status $emailAddress

                Action-GooglePerformSetMailboxForward -emailAddress $emailAddress -targetAddress $targetAddress
            }
        }
    }
    else
    {
        $emailAddress = (Helper-PromptString -prompt "Enter the email address of the mailbox" -allowEmpty $false)
        $targetAddress = (Helper-PromptString -prompt "Enter the email address to forward to" -allowEmpty $false)

        Action-GooglePerformSetMailboxForward -emailAddress $emailAddress -targetAddress $targetAddress
    }
}

function Action-GooglePerformSetMailboxForward([string]$emailAddress, [string]$targetAddress)
{
    try
    {
        GoogleHelper-SetMailboxForward -emailAddress $emailAddress -targetAddress $targetAddress
    }
    catch
    {
        Write-Host
        Write-Host -Object $_.ToString() -ForegroundColor Red
        Write-Host
        Write-Host -Object "Failed to set mailbox forward for $emailAddress to $targetAddress.  Possible causes include:" -ForegroundColor Red
        Write-Host -Object "   1) The email address $emailAddress does not correspond to a mailbox" -ForegroundColor Red
        Write-Host -Object "   2) The forwarding address is not allowed by Google.  It must adhere to the at least ONE of the following" -ForegroundColor Red
        Write-Host -Object "      2.1) The domain for $targetAddress must be added and verified within the G Suite account" -ForegroundColor Red
        Write-Host -Object "      2.2) The domain for $targetAddress must be a subdomain of one of the verified G Suite" -ForegroundColor Red
        Write-Host -Object "           domains within the G Suite account" -ForegroundColor Red

        throw
    }
}

######################################################################################################################################################
# Main menu -> Manage G Suite -> Remove email forwarding for mailbox
######################################################################################################################################################

function Action-GoogleRemoveMailboxForward()
{
    Write-Host

    do
    {
        $emailAddress = Read-Host -Prompt "Enter the email address of the mailbox"
    }
    while($emailAddress.Length -le 0)

    Write-Host -Object ("Removing mailbox forward for $emailAddress")
    GoogleHelper-RemoveMailboxForward -emailAddress $emailAddress
}

######################################################################################################################################################
# Main menu -> Manage G Suite -> Export users to CSV
######################################################################################################################################################

function Action-GoogleExportUsers()
{
    Write-Host

    do
    {
        $domainName = Read-Host -Prompt "Enter the G Suite domain to export"
    }
    while($domainName.Length -le 0)

    $users = GoogleHelper-GetDomainUsers -domainName $domainName
    if($users -ne $null)
    {
        $count = 0
        $csv = "DisplayName,FirstName,LastName,EmailAddress,Alias`r`n"
        $filename = Helper-GenerateRandomTempFilename -identifier "GoogleUsers"
        $file = New-Item -Path $filename -ItemType file -force -value $csv

        Write-Host -Object ("Creating MigrationWiz import file " + $filename)

        foreach($user in $users)
        {
            $count++
            Write-Progress -Activity ("Exporting G Suite user for " + $domainName + " (" + $count + ")") -Status $user.PrimaryEmail

            $csv = ""
            $csv += '"' + $user.Name.GivenName + " " + $user.Name.FamilyName + '"' + ','	# DisplayName
            $csv += '"' + $user.Name.GivenName + '"' + ','									# FirstName
            $csv += '"' + $user.Name.FamilyName + '"' + ','									# LastName
            $csv += '"' + $user.PrimaryEmail + '"' + ','								    # EmailAddress

            if($user.Aliases -ne $null -and $user.Aliases.Length -ge 1)
            {
                $csv += '"' + $user.Aliases[0] + '"' + ','									# Alias
            }

            Add-Content -Path $filename -Value $csv
        }
    }
    else
    {
        Write-Host
        Write-Host -Object "No users found within the domain specified.  Possible causes include:" -ForegroundColor Red
        Write-Host -Object "   1) no users" -ForegroundColor Red
        Write-Host -Object "   2) invalid domain specified" -ForegroundColor Red
        Write-Host -Object "   3) invalid credentials specified" -ForegroundColor Red
    }
}

######################################################################################################################################################
# Main menu -> Manage G Suite -> Revoke access token
######################################################################################################################################################

function Action-GoogleRevokeOAuth2AccessToken
{
    Write-Host

    GoogleHelper-RevokeOAuth2AccessToken
}

######################################################################################################################################################
# Main menu -> Manage Office 365
######################################################################################################################################################

function Menu-Office365Prompt()
{
    Write-Host
    Write-Host -Object ("Select an Office 365 task to perform:") -ForegroundColor Yellow
    Write-Host
    Write-Host -Object "0 - Export users to CSV"
    Write-Host -Object "1 - Export contacts to CSV"
    Write-Host -Object "2 - Disable password expiration on all users"
    Write-Host -Object "3 - Change UPN of all users (except admin account being used)"
    Write-Host -Object "4 - Set default domain"
    Write-Host -Object "5 - Set the same password for all users (except admin account being used)"
    Write-Host -Object "6 - Grant migration permissions to admin account"
    Write-Host -Object "7 - Set random password for all users and output to CSV (except admin account being used)"
    Write-Host -Object "8 - Create admin account for migration"
    Write-Host -Object "9 - Public Folders: Setup public folders from exported data (JSON file)"
    Write-Host -Object "10 - Export public folder data"
    Write-Host -Object "x - Back"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0-10 or x"
        if($result -eq "x")
        {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 10))
        {
            return [int]$result
        }
    }

    return $null
}

function Menu-Office365Loop()
{
    if(Office365Helper-GetCredentials)
    {
        # keep looping until specified to exit
        do
        {
            $action = Menu-Office365Prompt
            if($action -ne $null)
            {
                switch($action)
                {
                    0 # Export users to CSV
                    {
                        Action-Office365ExportUsers
                    }

                    1 # Export contacts to CSV
                    {
                        Action-Office365ExportContacts
                    }

                    2 # Disable password expiration on all users
                    {
                        Helper-LoadOffice365Module
                        Action-Office365DisableUserPasswords
                    }

                    3 # Change UPN of all users
                    {
                        Helper-LoadOffice365Module
                        Action-Office365ChangeUserPrincipalNames
                    }

                    4 # Set default domain
                    {
                        Helper-LoadOffice365Module
                        Action-Office365SetDefaultDomain
                    }

                    5 # Set the same password for all users
                    {
                        Helper-LoadOffice365Module
                        Action-Office365SetUserPasswords
                    }

                    6 # Grant migration permissions to admin account
                    {
                        Action-Office365GrantMigrationPermissions -userPrincipalName $script:o365Creds.UserName
                    }

                    7 # Generate random password for all users
                    {
                        Write-Host
                        if(Helper-PromptConfirmation -prompt "Are you sure you want to change all user passwords to something random?")
                        {
                            Helper-LoadOffice365Module
                            Action-Office365SetUserPasswordsRandom
                        }
                    }

                    8 # Create admin account for migration
                    {
                        Helper-LoadOffice365Module
                        Action-Office365CreateAdminAccount
                    }

                    9 # Public Folders: Setup public folders from exported JSON file
                    {
                        Helper-LoadOffice365Module
                        Action-Office365SetupPublicFolders
                    }

                    10 # Export public folder data
                    {
                        Action-ExportPublicFolderData -typeOfProject "O365"
                    }
                }
            }
            else
            {
                return
            }
        }
        while($true)
    }
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Export users to CSV
######################################################################################################################################################

function Action-Office365ExportUsers()
{
    try
    {
        if(Office365Helper-GetCredentials)
        {
            $count = 0
            $filename = Helper-GenerateRandomTempFilename -identifier "Office365Users"

            Write-Host
            Write-Host -Object ("Exporting Office 365 users to " + $filename)

            $csv = "AssistantName,AssistantPhone,City,Country,Company,Department,DisplayName,DisplayNamePrintable,ExtensionAttribute1,ExtensionAttribute2,ExtensionAttribute3,ExtensionAttribute4,ExtensionAttribute5,ExtensionAttribute6,ExtensionAttribute7,ExtensionAttribute8,ExtensionAttribute9,ExtensionAttribute10,ExtensionAttribute11,ExtensionAttribute12,ExtensionAttribute13,ExtensionAttribute14,ExtensionAttribute15,FaxPhone,FirstName,HomePhone,Initials,LastName,MobilePhone,Notes,ObjectGuid,Office,OtherFaxPhone,OtherHomePhone,OtherTelephone,PagerPhone,Telephone,PostalCode,RecipientType,State,StreetAddress,Title,WebPage,MicrosoftOnlineServicesId,SamAccountName,UserPrincipalName,WindowsEmailAddress,WindowsLiveId,Alias,Database,EmailAddresses,ExternalEmailAddress,ExchangeGuid,ForwardingSmtpAddress,HiddenFromAddressListsEnabled,LegacyExchangeDn,PrimarySmtpAddress`r`n"
            $file = New-Item -Path $filename -ItemType file -force -value $csv

            Office365Helper-ConnectRemotePowerShell

            $users = @(Get-User -ResultSize Unlimited)
            if($users -ne $null)
            {
                foreach($user in $users)
                {
                    $count++
                    Write-Progress -Activity ("Exporting Office 365 users (" + $count + "/" + $users.Length + ")") -Status $user.Name -PercentComplete ($count/$users.Length*100)

                    $mailbox = $null
                    $mailbox = Get-Mailbox -Identity $user.Guid.ToString() -ErrorAction SilentlyContinue

                    $mailUser = $null
                    $mailUser = Get-MailUser -Identity $user.Guid.ToString() -ErrorAction SilentlyContinue

                    $csv = ""
                    $csv += '"' + $user.AssistantName + '"' + ','									# AssistantName
                    $csv += '"' + $user.TelephoneAssistant + '"' + ','								# AssistantPhone
                    $csv += '"' + $user.City + '"' + ','											# City
                    $csv += '"' + $user.CountryOrRegion + '"' + ','									# Country
                    $csv += '"' + $user.Company + '"' + ','											# Company
                    $csv += '"' + $user.Department + '"' + ','										# Department
                    $csv += '"' + $user.DisplayName + '"' + ','										# DisplayName
                    $csv += '"' + $user.SimpleDisplayName + '"' + ','								# DisplayNamePrintable
                    $csv += '"' + $user.CustomAttribute1 + '"' + ','								# ExtensionAttribute1
                    $csv += '"' + $user.CustomAttribute2 + '"' + ','								# ExtensionAttribute2
                    $csv += '"' + $user.CustomAttribute3 + '"' + ','								# ExtensionAttribute3
                    $csv += '"' + $user.CustomAttribute4 + '"' + ','								# ExtensionAttribute4
                    $csv += '"' + $user.CustomAttribute5 + '"' + ','								# ExtensionAttribute5
                    $csv += '"' + $user.CustomAttribute6 + '"' + ','								# ExtensionAttribute6
                    $csv += '"' + $user.CustomAttribute7 + '"' + ','								# ExtensionAttribute7
                    $csv += '"' + $user.CustomAttribute8 + '"' + ','								# ExtensionAttribute8
                    $csv += '"' + $user.CustomAttribute9 + '"' + ','								# ExtensionAttribute9
                    $csv += '"' + $user.CustomAttribute10 + '"' + ','								# ExtensionAttribute10
                    $csv += '"' + $user.CustomAttribute11 + '"' + ','								# ExtensionAttribute11
                    $csv += '"' + $user.CustomAttribute12 + '"' + ','								# ExtensionAttribute12
                    $csv += '"' + $user.CustomAttribute13 + '"' + ','								# ExtensionAttribute13
                    $csv += '"' + $user.CustomAttribute14 + '"' + ','								# ExtensionAttribute14
                    $csv += '"' + $user.CustomAttribute15 + '"' + ','								# ExtensionAttribute15
                    $csv += '"' + $user.Fax + '"' + ','												# FaxPhone
                    $csv += '"' + $user.FirstName + '"' + ','										# FirstName
                    $csv += '"' + $user.HomePhone + '"' + ','										# HomePhone
                    $csv += '"' + $user.Initials + '"' + ','										# Initials
                    $csv += '"' + $user.LastName + '"' + ','										# LastName
                    $csv += '"' + $user.MobilePhone + '"' + ','										# MobilePhone
                    $csv += '"' + $user.Notes + '"' + ','											# Notes
                    $csv += '"' + $user.Guid.ToString() + '"' + ','									# ObjectGuid
                    $csv += '"' + $user.Office + '"' + ','											# Office
                    $csv += '"' + [string]::join(';', $user.OtherFax) + '"' + ','					# OtherFaxPhone
                    $csv += '"' + [string]::join(';', $user.OtherHomePhone) + '"' + ','				# OtherHomePhone
                    $csv += '"' + [string]::join(';', $user.OtherTelephone) + '"' + ','				# OtherTelephone
                    $csv += '"' + $user.Pager + '"' + ','											# PagerPhone
                    $csv += '"' + $user.Phone + '"' + ','											# Telephone
                    $csv += '"' + $user.PostalCode + '"' + ','										# PostalCode
                    $csv += '"' + $user.RecipientType + '"' + ','									# RecipientType
                    $csv += '"' + $user.StateOrProvince + '"' + ','									# State
                    $csv += '"' + $user.StreetAddress + '"' + ','									# StreetAddress
                    $csv += '"' + $user.Title + '"' + ','											# Title
                    $csv += '"' + $user.WebPage + '"' + ','											# WebPage

                    $csv += '"' + $user.MicrosoftOnlineServicesID + '"' + ','						# MicrosoftOnlineServicesId
                    $csv += '"' + $user.SamAccountName + '"' + ','									# SamAccountName
                    $csv += '"' + $user.UserPrincipalName + '"' + ','								# UserPrincipalName
                    $csv += '"' + $user.WindowsEmailAddress + '"' + ','								# WindowsEmailAddress
                    $csv += '"' + $user.WindowsLiveID + '"' + ','									# WindowsLiveId

                    if($mailbox -ne $null)
                    {
                        $csv += '"' + $mailbox.Alias + '"' + ','									# Alias
                        $csv += '"' + $mailbox.Database + '"' + ','									# Database
                        $csv += '"' + [string]::join(';', $mailbox.EmailAddresses) + '"' + ','		# EmailAddresses
                        $csv += ','																	# ExternalEmailAddress
                        $csv += '"' + $mailbox.ExchangeGuid + '"' + ','								# ExchangeGuid
                        $csv += '"' + $mailbox.ForwardingSmtpAddress + '"' + ','					# ForwardingSmtpAddress
                        $csv += '"' + $mailbox.HiddenFromAddressListsEnabled + '"' + ','			# HiddenFromAddressListsEnabled
                        $csv += '"' + $mailbox.LegacyExchangeDN + '"' + ','							# LegacyExchangeDN
                        $csv += '"' + $mailbox.PrimarySmtpAddress + '"'								# PrimarySmtpAddress
                    }
                    elseif($mailUser -ne $null)
                    {
                        $csv += '"' + $mailUser.Alias + '"' + ','									# Alias
                        $csv += ','																	# Database
                        $csv += '"' + [string]::join(';', $mailUser.EmailAddresses) + '"' + ','		# EmailAddresses
                        $csv += '"' + $mailUser.ExternalEmailAddress + '"' + ','					# ExternalEmailAddress
                        $csv += '"' + $mailUser.ExchangeGuid + '"' + ','							# ExchangeGuid
                        $csv += ','																	# ForwardingSmtpAddress
                        $csv += '"' + $mailUser.HiddenFromAddressListsEnabled + '"' + ','			# HiddenFromAddressListsEnabled
                        $csv += '"' + $mailUser.LegacyExchangeDN + '"' + ','						# LegacyExchangeDn
                        $csv += '"' + $mailUser.PrimarySmtpAddress + '"'							# PrimarySmtpAddress
                    }
                    else
                    {
                        $csv += ',,,,,,,,'
                    }

                    Add-Content -Path $filename -Value $csv
                }
            }
        }
    }
    finally
    {
        Get-PSSession | Remove-PSSession
    }
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Export contacts to CSV
######################################################################################################################################################

function Action-Office365ExportContacts()
{
    try
    {
        if(Office365Helper-GetCredentials)
        {
            $count = 0
            $filename = Helper-GenerateRandomTempFilename -identifier "Office365Contacts"

            Write-Host
            Write-Host -Object ("Exporting Office 365 contacts to " + $filename)

            $csv = "AssistantName,AssistantPhone,City,Country,Company,Department,DisplayName,DisplayNamePrintable,ExtensionAttribute1,ExtensionAttribute2,ExtensionAttribute3,ExtensionAttribute4,ExtensionAttribute5,ExtensionAttribute6,ExtensionAttribute7,ExtensionAttribute8,ExtensionAttribute9,ExtensionAttribute10,ExtensionAttribute11,ExtensionAttribute12,ExtensionAttribute13,ExtensionAttribute14,ExtensionAttribute15,FaxPhone,FirstName,HomePhone,Initials,LastName,MobilePhone,Notes,ObjectGuid,Office,OtherFaxPhone,OtherHomePhone,OtherTelephone,PagerPhone,Telephone,PostalCode,RecipientType,State,StreetAddress,Title,WebPage,MicrosoftOnlineServicesId,SamAccountName,UserPrincipalName,WindowsEmailAddress,WindowsLiveId,Alias,Database,EmailAddresses,ExternalEmailAddress,ExchangeGuid,ForwardingSmtpAddress,HiddenFromAddressListsEnabled,LegacyExchangeDn,PrimarySmtpAddress`r`n"
            $file = New-Item -Path $filename -ItemType file -force -value $csv

            Office365Helper-ConnectRemotePowerShell

            $contacts = @(Get-Contact -ResultSize Unlimited)
            if($contacts -ne $null)
            {
                foreach($contact in $contacts)
                {
                    $count++
                    Write-Progress -Activity ("Exporting Office 365 contacts (" + $count + "/" + $contacts.Length + ")") -Status $contact.Name -PercentComplete ($count/$contacts.Length*100)

                    $mailContact = $null
                    $mailContact = Get-MailContact -Identity $contact.Guid.ToString() -ErrorAction SilentlyContinue

                    $csv = ""
                    $csv += '"' + $contact.AssistantName + '"' + ','								# AssistantName
                    $csv += '"' + $contact.TelephoneAssistant + '"' + ','							# AssistantPhone
                    $csv += '"' + $contact.City + '"' + ','											# City
                    $csv += '"' + $contact.CountryOrRegion + '"' + ','								# Country
                    $csv += '"' + $contact.Company + '"' + ','										# Company
                    $csv += '"' + $contact.Department + '"' + ','									# Department
                    $csv += '"' + $contact.DisplayName + '"' + ','									# DisplayName
                    $csv += '"' + $contact.SimpleDisplayName + '"' + ','							# DisplayNamePrintable
                    $csv += '"' + $contact.CustomAttribute1 + '"' + ','								# ExtensionAttribute1
                    $csv += '"' + $contact.CustomAttribute2 + '"' + ','								# ExtensionAttribute2
                    $csv += '"' + $contact.CustomAttribute3 + '"' + ','								# ExtensionAttribute3
                    $csv += '"' + $contact.CustomAttribute4 + '"' + ','								# ExtensionAttribute4
                    $csv += '"' + $contact.CustomAttribute5 + '"' + ','								# ExtensionAttribute5
                    $csv += '"' + $contact.CustomAttribute6 + '"' + ','								# ExtensionAttribute6
                    $csv += '"' + $contact.CustomAttribute7 + '"' + ','								# ExtensionAttribute7
                    $csv += '"' + $contact.CustomAttribute8 + '"' + ','								# ExtensionAttribute8
                    $csv += '"' + $contact.CustomAttribute9 + '"' + ','								# ExtensionAttribute9
                    $csv += '"' + $contact.CustomAttribute10 + '"' + ','							# ExtensionAttribute10
                    $csv += '"' + $contact.CustomAttribute11 + '"' + ','							# ExtensionAttribute11
                    $csv += '"' + $contact.CustomAttribute12 + '"' + ','							# ExtensionAttribute12
                    $csv += '"' + $contact.CustomAttribute13 + '"' + ','							# ExtensionAttribute13
                    $csv += '"' + $contact.CustomAttribute14 + '"' + ','							# ExtensionAttribute14
                    $csv += '"' + $contact.CustomAttribute15 + '"' + ','							# ExtensionAttribute15
                    $csv += '"' + $contact.Fax + '"' + ','											# FaxPhone
                    $csv += '"' + $contact.FirstName + '"' + ','									# FirstName
                    $csv += '"' + $contact.HomePhone + '"' + ','									# HomePhone
                    $csv += '"' + $contact.Initials + '"' + ','										# Initials
                    $csv += '"' + $contact.LastName + '"' + ','										# LastName
                    $csv += '"' + $contact.MobilePhone + '"' + ','									# MobilePhone
                    $csv += '"' + $contact.Notes + '"' + ','										# Notes
                    $csv += '"' + $contact.Guid.ToString() + '"' + ','								# ObjectGuid
                    $csv += '"' + $contact.Office + '"' + ','										# Office
                    $csv += '"' + [string]::join(';', $contact.OtherFax) + '"' + ','				# OtherFaxPhone
                    $csv += '"' + [string]::join(';', $contact.OtherHomePhone) + '"' + ','			# OtherHomePhone
                    $csv += '"' + [string]::join(';', $contact.OtherTelephone) + '"' + ','			# OtherTelephone
                    $csv += '"' + $contact.Pager + '"' + ','										# PagerPhone
                    $csv += '"' + $contact.Phone + '"' + ','										# Telephone
                    $csv += '"' + $contact.PostalCode + '"' + ','									# PostalCode
                    $csv += '"' + $contact.RecipientType + '"' + ','								# RecipientType
                    $csv += '"' + $contact.StateOrProvince + '"' + ','								# State
                    $csv += '"' + $contact.StreetAddress + '"' + ','								# StreetAddress
                    $csv += '"' + $contact.Title + '"' + ','										# Title
                    $csv += '"' + $contact.WebPage + '"' + ','										# WebPage

                    $csv += ','																		# MicrosoftOnlineServicesId
                    $csv += ','																		# SamAccountName
                    $csv += ','																		# UserPrincipalName
                    $csv += ','																		# WindowsEmailAddress
                    $csv += ','																		# WindowsLiveId

                    if($mailContact -ne $null)
                    {
                        $csv += '"' + $mailContact.Alias + '"' + ','								# Alias
                        $csv += ','																	# Database
                        $csv += '"' + [string]::join(';', $mailContact.EmailAddresses) + '"' + ','	# EmailAddresses
                        $csv += '"' + $mailContact.ExternalEmailAddress + '"' + ','					# ExternalEmailAddress
                        $csv += '"' + $mailContact.ExchangeGuid + '"' + ','							# ExchangeGuid
                        $csv += ','																	# ForwardingSmtpAddress
                        $csv += '"' + $mailContact.HiddenFromAddressListsEnabled + '"' + ','		# HiddenFromAddressListsEnabled
                        $csv += '"' + $mailContact.LegacyExchangeDN + '"' + ','						# LegacyExchangeDn
                        $csv += '"' + $mailContact.PrimarySmtpAddress + '"'							# PrimarySmtpAddress
                    }
                    else
                    {
                        $csv += ',,,,,,,,'
                    }

                    Add-Content -Path $filename -Value $csv
                }
            }
        }
    }
    finally
    {
        Get-PSSession | Remove-PSSession
    }
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Disable password expiration on all users
######################################################################################################################################################

function Action-Office365DisableUserPasswords
{
    Connect-MsolService -Credential (Office365Helper-GetCredentials)

    Write-Host
    Write-Host -Object ("Disabling password expiration for all Office 365 users ...")

    $users = @(Get-MsolUser -All)
    if($users -ne $null)
    {
        foreach($user in $users)
        {
            $count++
            Write-Progress -Activity ("Disabling Office 365 user password expiration (" + $count + "/" + $users.Length + ")") -Status $user.DisplayName -PercentComplete ($count/$users.Length*100)

            $result = Set-MsolUser -ObjectId $user.ObjectId.ToString() -PasswordNeverExpires $true
        }
    }

}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Change UPN of all users
######################################################################################################################################################

function Action-Office365ChangeUserPrincipalNames
{
    $count = 0
    $domainName = $null

    Connect-MsolService -Credential (Office365Helper-GetCredentials)
    $adminUpn = $script:o365Creds.UserName

    Write-Host
    $domainName = (Helper-PromptString -prompt "Enter the domain name to change all users to" -allowEmpty $false)
    Write-Host -Object ("Changing user principal name of all Office 365 users except " + $adminUpn)

    $users = @(Get-MsolUser -All)
    if($users -ne $null)
    {
        foreach($user in $users)
        {
            $count++
            Write-Progress -Activity ("Setting Office 365 user principal name (" + $count + "/" + $users.Length + ")") -Status $user.DisplayName -PercentComplete ($count/$users.Length*100)

            if($user.UserPrincipalName.ToLower() -ne $adminUpn.ToLower())
            {
                $newUpn = ($user.UserPrincipalName.Split("@")[0] + "@" + $domainName)
                Set-MsolUserPrincipalName -ObjectId $user.ObjectId -NewUserPrincipalName $newUpn
            }
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Set default domain
######################################################################################################################################################

function Action-Office365SetDefaultDomain
{
    $domainName = $null

    Connect-MsolService -Credential (Office365Helper-GetCredentials)

    Write-Host
    $domainName = (Helper-PromptString -prompt "Enter the domain name to set as default" -allowEmpty $false)
    Write-Host -Object ("Setting default domain to " + $domainName)

    Set-MsolDomain -Name $domainName –IsDefault
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Set the same password for all users
######################################################################################################################################################

function Action-Office365SetUserPasswords
{
    $count = 0
    $password = $null

    Connect-MsolService -Credential (Office365Helper-GetCredentials)
    $adminUpn = $script:o365Creds.UserName

    Write-Host
    $password = (Helper-PromptString -prompt "Enter the new password to set for all users" -allowEmpty $false)
    $forceChange = (Helper-PromptConfirmation -prompt "Would you like to user to be forced to change the password on first login?")
    Write-Host -Object ("Changing all Office 365 user passwords except " + $adminUpn)

    $users = @(Get-MsolUser -All)
    if($users -ne $null)
    {
        foreach($user in $users)
        {
            $count++
            Write-Progress -Activity ("Setting Office 365 user password (" + $count + "/" + $users.Length + ")") -Status $user.DisplayName -PercentComplete ($count/$users.Length*100)

            if($user.UserPrincipalName.ToLower() -ne $adminUpn.ToLower())
            {
                $result = Set-MsolUserPassword -ObjectId $user.ObjectId.ToString() -NewPassword $password –ForceChangePassword $forceChange
            }
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Grant migration permissions to admin account
######################################################################################################################################################

function Action-Office365GrantMigrationPermissions([string]$userPrincipalName)
{
    try
    {
        $count = 0

        Office365Helper-ConnectRemotePowerShell

        $adminUser = (Get-User -ResultSize Unlimited | Where-Object -FilterScript { $_.UserPrincipalName.ToLower() -eq $userPrincipalName.ToLower() })
        if($adminUser -ne $null)
        {
            Write-Host
            Write-Host -Object ("Granting Office 365 delegation permissions for all users to " + $adminUser.UserPrincipalName)

            $users = @(Get-Mailbox -ResultSize Unlimited)
            if($users -ne $null)
            {
                foreach($user in $users)
                {
                    $count++
                    Write-Progress -Activity ("Granting Office 365 migration permissions (" + $count + "/" + $users.Length + ")") -Status $user.Name -PercentComplete ($count/$users.Length*100)

                    $result = Add-MailboxPermission -Identity $user.Guid.ToString() -AccessRights FullAccess -User $adminUser.Identity -Automapping $false -WarningAction SilentlyContinue
                }
            }

            Write-Host -Object ("Granting Office 365 impersonation permissions for all users to " + $adminUser.UserPrincipalName)
            Enable-OrganizationCustomization -ErrorAction SilentlyContinue
            $result = New-ManagementRoleAssignment -Role ApplicationImpersonation -User $adminUser.Identity -WarningAction SilentlyContinue
        }
    }
    finally
    {
        Get-PSSession | Remove-PSSession
    }
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Generate random password for all users
######################################################################################################################################################

function Action-Office365SetUserPasswordsRandom
{
    $count = 0
    $filename = Helper-GenerateRandomTempFilename -identifier "Office365UserPasswords"

    Connect-MsolService -Credential (Office365Helper-GetCredentials)
    $adminUpn = $script:o365Creds.UserName

    Write-Host
    $forceChange = (Helper-PromptConfirmation -prompt "Would you like to user to be forced to change the password on first login?")

    Write-Host
    Write-Host -Object ("Passwords will be saved to $filename")
    Write-Host -Object ("Changing all Office 365 user passwords to something random except $adminUpn")

    $csv = "UserPrincipalName,Password`r`n"
    $file = New-Item -Path $filename -ItemType file -force -value $csv

    $users = @(Get-MsolUser -All)
    if($users -ne $null)
    {
        foreach($user in $users)
        {
            $count++
            Write-Progress -Activity ("Setting Office 365 user password (" + $count + "/" + $users.Length + ")") -Status $user.DisplayName -PercentComplete ($count/$users.Length*100)

            if($user.UserPrincipalName.ToLower() -ne $adminUpn.ToLower())
            {
                $userPrincipalName = $user.UserPrincipalName
                $password = Helper-GeneratePassword
                Helper-WriteDebug -line ("UPN = $userPrincipalName, Password = $password")

                $result = Set-MsolUserPassword -ObjectId $user.ObjectId.ToString() -NewPassword $password –ForceChangePassword $forceChange

                $csv = ""
                $csv += '"' + $userPrincipalName + '"' + ','		# UserPrincipalName
                $csv += '"' + $password + '"'						# Password

                Add-Content -Path $filename -Value $csv
            }
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Create admin account
######################################################################################################################################################

function Action-Office365CreateAdminAccount
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword')]
    param()

    try
    {
        Write-Host

        Office365Helper-ConnectRemotePowerShell
        Connect-MsolService -Credential (Office365Helper-GetCredentials)

        $domains = @(Get-MsolDomain)
        if($domains -ne $null -or $domains.Count -le 0)
        {
            foreach($domain in $domains)
            {
                Helper-WriteDebug -line ("Domain = $domain")
                if($domain.Name.ToLower().EndsWith(".onmicrosoft.com") -and $domain.Name.Split('.').Length -eq 3)
                {
                    $adminEmail = $script:migrationAdminName + "@" + $domain.Name
                    Helper-WriteDebug -line ("Admin Email Address = $adminEmail")

                    $adminMailbox = (Get-Mailbox -Identity $adminEmail -ErrorAction SilentlyContinue)
                    if($adminMailbox -eq $null)
                    {
                        Write-Host -Object ("Creating migration admin account $adminEmail")

                        $adminPassword = Helper-GeneratePassword
                        $adminMailbox = New-Mailbox -Name $script:migrationAdminName -MicrosoftOnlineServicesID $adminEmail -Password (ConvertTo-SecureString -String $adminPassword -AsPlainText -Force) -RemotePowerShellEnabled $true -WarningAction SilentlyContinue

                        Write-Host -Object ("Admin UserName = $adminEmail") -ForegroundColor Yellow
                        Write-Host -Object ("Admin Password = $adminPassword") -ForegroundColor Yellow
                    }
                    else
                    {
                        Write-Host -Object ("The migration admin account $adminEmail already exists")
                    }

                    $adminMsolUser = Office365Helper-WaitForReplication -userPrincipalName $adminEmail
                    if($adminMsolUser.Licenses -eq $null -or $adminMsolUser.Licenses.Count -le 0)
                    {
                        $sku = Office365Helper-PromptSku
                        if($sku -ne $null)
                        {
                            $usageLocation = (Helper-PromptString -prompt "Enter the two letter country location of the mailbox (i.e US)" -allowEmpty $false)

                            Write-Host -Object ("Assigning license to mailbox")

                            $licenseOption = New-MsolLicenseOptions -AccountSkuId $sku.AccountSkuId
                            Set-MsolUser -UserPrincipalName $adminEmail -UsageLocation $usageLocation
                            Set-MsolUserLicense -UserPrincipalName $adminEmail -AddLicenses $sku.AccountSkuId -LicenseOptions $licenseOption
                        }
                        else
                        {
                            throw "No Office 365 licenses found"
                        }
                    }
                    else
                    {
                        Write-Host -Object "Office 365 license is already assigned"
                    }

                    Action-Office365GrantMigrationPermissions -userPrincipalName $adminEmail

                    break
                }
                else
                {
                    Helper-WriteDebug -line ("Found domain " + $domain.Name)
                }
            }
        }
        else
        {
            Write-Host -Object "No domain names registered within account"
        }
    }
    finally
    {
        Get-PSSession | Remove-PSSession
    }
}

######################################################################################################################################################
# Main menu -> Manage Office 365 -> Setup public folders from exported JSON file
######################################################################################################################################################

function Action-Office365SetupPublicFolders
{
    try
    {
        Write-Host

        Office365Helper-ConnectRemotePowerShell
        $liveCred = (Office365Helper-GetCredentials)
        Connect-MsolService -Credential $liveCred

        # Retrieve the mapping file
        if ($script:allExchangeServerPublicFoldersPerMailboxFileName -ne $null)
        {
            $mappingFile = (Helper-PromptString -prompt "Enter the location of the JSON file[$script:allExchangeServerPublicFoldersPerMailboxFileName]" -allowEmpty $true)
            if ($mappingFile -eq $null)
            {
                $mappingFile = $script:allExchangeServerPublicFoldersPerMailboxFileName
            }
        } 
        else
        {
            $mappingFile = (Helper-PromptString -prompt "Enter the location of the JSON file" -allowEmpty $false)
        }

        # Retrieve all objects from the mapping file
        Write-Host -Object "Converting mapping file to objects ..."
        $publicFolderMailboxesRaw = Get-Content $mappingFile | ConvertFrom-Json
        $publicFolderMailboxes = @()

        # Convert from raw to PublicFolderMailboxData
        foreach($publicFolderMailboxRaw in $publicFolderMailboxesRaw)
        {
            $publicFolderMailboxData = New-Object -TypeName BitTitan.ExchangeTools.PowerShell.Data.PublicFolderMailboxData

            # Set the mailbox data
            $publicFolderMailboxData.MailboxName = $publicFolderMailboxRaw.MailboxName
            $publicFolderMailboxData.PublicFolderList = New-Object System.Collections.Generic.List[BitTitan.ExchangeTools.PowerShell.Data.PublicFolderData]

            # Get all of the public folder data
            foreach($publicFolderRaw in $publicFolderMailboxRaw.PublicFolderList)
            {
                # Create the public folder data object
                $publicFolderData = New-Object -TypeName BitTitan.ExchangeTools.PowerShell.Data.PublicFolderData

                # Set the public folder data
                $publicFolderData.FolderClass = $publicFolderRaw.FolderClass
                $publicFolderData.FolderSize = $publicFolderRaw.FolderSize
                $publicFolderData.Identity = $publicFolderRaw.Identity
                $publicFolderData.ParentPath = $publicFolderRaw.ParentPath
                $publicFolderData.IsValid = $publicFolderRaw.IsValid
                $publicFolderData.ItemCount = $publicFolderRaw.ItemCount
                $publicFolderData.MailEnabled = $publicFolderRaw.MailEnabled
                $publicFolderData.Name = $publicFolderRaw.Name
                
                # Get the email addresses
                foreach($emailAddressRaw in $publicFolderRaw.EmailAddresses)
                {
                    # Create the email address object
                    $emailAddress = New-Object -TypeName BitTitan.ExchangeTools.PowerShell.Data.PublicFolderEmailAddress

                    # Set the fields
                    $emailAddress.SmtpAddress = $emailAddressRaw.SmtpAddress
                    $emailAddress.AddressString = $emailAddressRaw.AddressString
                    $emailAddress.ProxyAddressString = $emailAddressRaw.ProxyAddressString
                    $emailAddress.PrefixString = $emailAddressRaw.PrefixString
                    $emailAddress.IsPrimaryAddress = $emailAddressRaw.IsPrimaryAddress

                    # Add the element
                    $publicFolderData.EmailAddresses += $emailAddress
                }

                # Add the element to the list
                $publicFolderMailboxData.PublicFolderList.Add($publicFolderData)
            }

            # Add to the set of mailboxes
            $publicFolderMailboxes += $publicFolderMailboxData
        }

        # Write progress
        Write-Host -Object "Creating public folder mailboxes ..."

        # Select the mailboxes and create them
        $mailboxes = Select-PublicFolderData -Format UniqueMailboxes -PublicFolderMailboxData $publicFolderMailboxes
        foreach($mailboxName in $mailboxes)
        {
            # Check if it already exists
            $mailbox = (Get-Mailbox -Identity "$mailboxName" -ErrorAction SilentlyContinue -PublicFolder)

            # Create it when it does not exist
            if ($mailbox -eq $null)
            {
                New-Mailbox -Name "$mailboxName" -PublicFolder | Out-Null
                Write-Host -Object "  Created public folder mailbox $mailboxName"
            }
        }

        # Select public folders mailbox mappings and create the public folders in the mailboxes
        $mappings = Select-PublicFolderData -Format FoldersWithMailboxes -PublicFolderMailboxData $publicFolderMailboxes
        $mappings = $mappings | Sort-Object -Property Identity
        Write-Host -Object "Attempting to retrieve or creating $($mappings.Length) public folders ..."
        foreach($mapping in $mappings)
        {
            # Check if the folder exists
            $existingFolder = Get-PublicFolder -Identity $mapping.Identity -ErrorAction SilentlyContinue

            # Create the folder
            if ($existingFolder -eq $null)
            {
                New-PublicFolder -Name $mapping.Name -Path $mapping.ParentPath -Mailbox $mapping.MailboxName | Out-Null
                Write-Host -Object "Created public folder:  $($mapping.Identity)"
                
                # Assign owner permissions for the newly created folder to the current user
                if ($liveCred -ne $null)
                {
                    Add-PublicFolderClientPermission -Identity $mapping.Identity -User $liveCred.UserName -AccessRights Owner -ErrorAction SilentlyContinue | Out-Null
                }
            }
            else
            {
                Write-Host -Object "Existing public folder: $($mapping.Identity)"
            }
        }

        # Set the mail-enabled flag
        Write-Host -Object "Setting mail-enabled status on public folders"

        # Select mail-enabled public folders and set the mail-enabled flag
        $foldersToMailEnable = Select-PublicFolderData -Format MailEnabledFolderIdentities -PublicFolderMailboxData $publicFolderMailboxes
        $foldersToMailEnable | Enable-MailPublicFolder | Out-Null

        # Select the email addresses and create
        $foldersAndEmailAddresses = Select-PublicFolderData -Format FolderEmailAddresses -PublicFolderMailboxData $publicFolderMailboxes
        foreach($folder in $foldersAndEmailAddresses)
        {
            # Get the mail-enabled public folder
            $mailFolder = Get-MailPublicFolder -Identity $folder.Identity
            if ($mailFolder -eq $null) {
                Write-Warning -Message "Cannot find mail-enabled folder to write email addresses: $($folder.Identity)"
                continue
            }

            # Select all of the current addresses
            $destinationSmtpAddresses = $mailFolder.EmailAddresses | Where-Object -FilterScript { $_.SmtpAddress -ne $null } | Select-Object -Property SmtpAddress

            # Go through the email addresses
            foreach($emailAddress in $folder.EmailAddresses)
            {
                # Skip the primary address
                if ($destinationSmtpAddresses -contains $emailAddress.SmtpAddress) {
                    continue
                }

                # Add the SMTP address to the list of email addresses
                $mailFolder.EmailAddresses += $emailAddress.SmtpAddress
            }

            # Update the email addresses
            Set-MailPublicFolder -Identity $folder.Identity -EmailAddresses $mailFolder.EmailAddresses | Out-Null
        }

        Write-Host -Object "Public folder configuration complete"
    }
    finally
    {
        Get-PSSession | Remove-PSSession
    }
}
######################################################################################################################################################
# Main menu -> Manage local Exchange Server
######################################################################################################################################################

function Menu-ExchangeServerPrompt()
{
    Write-Host
    Write-Host -Object ("Select an local Exchange Server task to perform:") -ForegroundColor Yellow
    Write-Host
    Write-Host -Object "0 - Check alias for invalid characters"
    Write-Host -Object "1 - Identify user principal names that do not match the primary email addresses"
    Write-Host -Object "2 - Set user principal name to the same as the email address"
    Write-Host -Object "3 - Create contact forward and set mailbox forwarding"
    Write-Host -Object "4 - Create MigrationWiz mailbox import file"
    Write-Host -Object "5 - Create contact forwards for all mailboxes only"
    Write-Host -Object "6 - Export public folder data"
    Write-Host -Object "x - Back"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0-6 or x"
        if($result -eq "x")
        {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 6))
        {
            return [int]$result
        }
    }

    return $null
}

function Menu-ExchangeServerLoop()
{
    # keep looping until specified to exit
    do
    {
        $action = Menu-ExchangeServerPrompt
        if($action -ne $null)
        {
            switch($action)
            {
                0 # Check alias for invalid characters
                {
                    Action-ExchangeServerCheckAlias
                }

                1 # Identify user principal names that do not match the primary email addresses
                {
                    Action-ExchangeServerCompareEmailAddressAndUpn
                }

                2 # Set user principal name to the same as the email address
                {
                    Action-ExchangeServerSetUpnAsEmailAddress
                }

                3 # Create contact forward and set mailbox forwarding
                {
                    Action-ExchangeServerMailboxToContact
                }

                4 # Create MigrationWiz mailbox import file
                {
                    Action-ExchangeServerExportMailboxes
                }

                5 # Create contact forwards for all mailboxes only
                {
                    Action-ExchangeServerCreateContactForwardsOnly
                }
                6 # Export public folder data
                {
                    Action-ExportPublicFolderData -typeOfProject "Local"
                }
            }
        }
        else
        {
            return
        }
    }
    while($true)
}

######################################################################################################################################################
# Main menu -> Manage local Exchange Server -> Check alias for invalid characters
######################################################################################################################################################

function Action-ExchangeServerCheckAlias()
{
    $count = 0
    $filename = Helper-GenerateRandomTempFilename -identifier "ExchangeServerInvalidMailNickNames"
    $validCharacters = "^[ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstyvwxyz01234567890!#$%&'*+-/=?.^_`{|}~]+$"

    Write-Host
    Write-Host -Object ("Exporting invalid aliases to " + $filename)

    $csv = "Path,Alias`r`n"
    $file = New-Item -Path $filename -ItemType file -force -value $csv

    $filter = "(mailNickname=*)"
    $properties = @("mailnickname")
    $objects = @(ExchangeServerHelper-AdSearch -searchFilter $filter -searchProperties $properties)
    if($object -ne $null -and $objects.Length -ge 1)
    {
        foreach($object in $objects)
        {
            $ldapPath = $object.Path
            $mailNickname = $object.Properties["mailnickname"][0].ToString()

            $count++
            Write-Progress -Activity ("Checking alias value on object (" + $count + "/" + $objects.Length + ")") -Status $mailNickname -PercentComplete ($count/$objects.Length*100)

            Helper-WriteDebug -line ("MailNickname = " + $mailNickname)

            if($mailNickname -notmatch $validCharacters)
            {
                $csv = ""
                $csv += '"' + $ldapPath + '"' + ','		# Path
                $csv += '"' + $mailNickname + '"'		# Alias

                Add-Content -Path $filename -Value $csv
            }
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage local Exchange Server -> Identify user principal names that do not match the primary email addresses
######################################################################################################################################################

function Action-ExchangeServerCompareEmailAddressAndUpn()
{
    $count = 0
    $filename = Helper-GenerateRandomTempFilename -identifier "ExchangeServerEmailUpnComparison"

    Write-Host
    Write-Host -Object ("Exporting non-matching email address and UPNs to " + $filename)

    $csv = "Path,WindowsEmailAddress,UserPrincipalName`r`n"
    $file = New-Item -Path $filename -ItemType file -force -value $csv

    $filter = "(&(objectCategory=person)(objectClass=user)(mail=*))"
    $properties = @("mail", "userPrincipalName")
    $objects = @(ExchangeServerHelper-AdSearch -searchFilter $filter -searchProperties $properties)
    if($objects -ne $null -and $objects.Length -ge 1)
    {
        foreach($object in $objects)
        {
            $ldapPath = $object.Path
            $mail = $object.Properties["mail"][0].ToString()
            $userPrincipalName = $object.Properties["userprincipalname"][0].ToString()

            $count++
            Write-Progress -Activity ("Checking user principal name (" + $count + "/" + $objects.Length + ")") -Status $mail -PercentComplete ($count/$objects.Length*100)

            Helper-WriteDebug -line ("Mail = " + $mail + ", UPN = " + $userPrincipalName)

            if($mail.ToLower() -notmatch $userPrincipalName.ToLower())
            {
                $csv = ""
                $csv += '"' + $ldapPath + '"' + ','		# Path
                $csv += '"' + $mail + '"' + ','			# WindowsEmailAddress
                $csv += '"' + $userPrincipalName + '"'	# UserPrincipalName

                Add-Content -Path $filename -Value $csv
            }
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage local Exchange Server -> Set user principal name to the same as the email address
######################################################################################################################################################

function Action-ExchangeServerSetUpnAsEmailAddress()
{
    $count = 0

    Write-Host
    $continue = (Helper-PromptConfirmation -prompt "Are you sure you wish to continue?")
    if($continue)
    {
        Write-Host
        Write-Host -Object ("Setting user UPNs to match the email address ...")

        $filter = "(&(objectCategory=person)(objectClass=user)(mail=*))"
        $properties = @("mail", "userPrincipalName")
        $objects = @(ExchangeServerHelper-AdSearch -searchFilter $filter -searchProperties $properties)
        if($objects -ne $null -and $objects.Length -ge 1)
        {
            foreach($object in $objects)
            {
                $ldapPath = $object.Path
                $mail = $object.Properties["mail"][0].ToString()
                #$userPrincipalName = $object.Properties["userprincipalname"][0].ToString()

                $count++
                Write-Progress -Activity ("Setting user principal name (" + $count + "/" + $objects.Length + ")") -Status $mail -PercentComplete ($count/$objects.Length*100)

                #Helper-WriteDebug -line ("Mail = " + $mail + ", UPN = " + $userPrincipalName)

                $object = [ADSI]$ldapPath
                $object.Put("userPrincipalName", $mail)
                $object.SetInfo()
            }
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage local Exchange Server -> Create contact forwards for all mailboxes only
######################################################################################################################################################

function Action-ExchangeServerCreateContactForwardsOnly()
{
    $count = 0

    Write-Host
    $targetDomain = (Helper-PromptString -prompt "Enter the non-athoritative domain to redirect mail to (i.e. example.onmicrosoft.com)" -allowEmpty $false)

    Write-Host
    Write-Host -Object ("Creating contact forward objects for all mailboxes to $targetDomain ...")

    $filter = "(&(objectCategory=person)(objectClass=user)(mailNickname=*)(homeMdb=*)(mail=*))"
    $properties = @("objectClass", "mail")
    $objects = @(ExchangeServerHelper-AdSearch -searchFilter $filter -searchProperties $properties)
    if($objects -ne $null -and $objects.Length -ge 1)
    {
        foreach($object in $objects)
        {
            Helper-WriteDebug -line ("Found user")
            $user = [ADSI]$object.Path

            $emailAddress = $user.Properties["mail"][0]
            $targetAddress = ($emailAddress.Split("@")[0] + "@" + $targetDomain)

            ExchangeServerHelper-CreateContactForward -user $user -targetAddress $targetAddress -setForward $false
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage local Exchange Server -> Create contact forward and set mailbox forwarding
######################################################################################################################################################

function Action-ExchangeServerMailboxToContact()
{
    $count = 0

    Write-Host

    $importConfirm = (Helper-PromptConfirmation -prompt "Would you like to import a list from a file?")
    if($importConfirm)
    {
        $importFilename = (Helper-PromptString -prompt "Enter the full path to import file (Press enter to create one)" -allowEmpty $true)
        if($importFilename -eq "")
        {
            # create new import file
            $importFilename = Helper-GenerateRandomTempFilename -identifier "ExchangeServerForwardingImport"
            $csv = "Email Address,Forwarding Address`r`n"
            $file = New-Item -Path $importFilename -ItemType file -force -value $csv

            # open file for editing
            Start-Process -FilePath $importFilename

            do
            {
                $importConfirm = (Helper-PromptConfirmation -prompt "Are you done editing the import file?")
            }
            while(-not $importConfirm)
        }

        # read csv file
        $users = Import-Csv -Path $importFilename
        foreach($user in $users)
        {
            $emailAddress = $user.'Email Address'
            $targetAddress = $user.'Forwarding Address'

            if($emailAddress -ne $null -and $emailAddress -ne "" -and $targetAddress -ne $null -and $targetAddress -ne "")
            {
                $count++
                Write-Progress -Activity ("Converting mailbox to contact forward (" + $count + ")") -Status $emailAddress

                Action-ExchangeServerPerformMailboxToContact -emailAddress $emailAddress -targetAddress $targetAddress
            }
        }
    }
    else
    {
        $emailAddress = (Helper-PromptString -prompt "Enter the primary email address of the mailbox to convert" -allowEmpty $false)
        $targetAddress = (Helper-PromptString -prompt "Enter the non-authoritative email address to redirect mail to" -allowEmpty $false)

        Action-ExchangeServerPerformMailboxToContact -emailAddress $emailAddress -targetAddress $targetAddress
    }
}

function Action-ExchangeServerPerformMailboxToContact([string]$emailAddress, [string]$targetAddress)
{
    $filter = "(|(mail=$emailAddress)(proxyAddresses=smtp:$emailAddress))"
    $properties = @("objectClass")
    $objects = @(ExchangeServerHelper-AdSearch -searchFilter $filter -searchProperties $properties)
    if($objects -ne $null -and $objects.Length -ge 1)
    {
        if($objects.Length -eq 1)
        {
            $object = $objects[0]
            $objectClass = ($object.Properties["objectclass"] -join ";")
            Helper-WriteDebug -line ("Object-Class: " + $objectClass)

            if($objectClass.ToLower() -eq "top;person;organizationalperson;user")
            {
                Helper-WriteDebug -line ("Found user")

                $user = $null
                $user = [ADSI]$object.Path

                ExchangeServerHelper-CreateContactForward -user $user -targetAddress $targetAddress -setForward $true
            }
            else
            {
                Write-Host -Object "An non-user object of type $objectClass was found with that email address" -ForegroundColor Yellow
            }
        }
        else
        {
            Write-Host -Object "More than one object found with the email address of $emailAddress.  Clean up your AD then try again." -ForegroundColor Yellow
        }
    }
    else
    {
        Write-Host -Object "No object found with that email address" -ForegroundColor Yellow
    }
}

######################################################################################################################################################
# Main menu -> Manage local Exchange Server -> Create MigrationWiz mailbox import file
######################################################################################################################################################

function Action-ExchangeServerExportMailboxes()
{
    $count = 0
    $filename = Helper-GenerateRandomTempFilename -identifier "ExchangeServerMailboxes"

    Write-Host
    Write-Host -Object ("Creating MigrationWiz import file " + $filename)

    $csv = "Source Email,Source UserName,Source Password,Destination Email,Destination UserName,Destination Password`r`n"
    $file = New-Item -Path $filename -ItemType file -force -value $csv

    $filter = "(&(objectCategory=person)(objectClass=user)(mailNickname=*)(homeMdb=*)(mail=*))"
    $properties = @("mail")
    $objects = @(ExchangeServerHelper-AdSearch -searchFilter $filter -searchProperties $properties)
    if($objects -ne $null -and $objects.Length -ge 1)
    {
        foreach($object in $objects)
        {
            $ldapPath = $object.Path
            $mail = $object.Properties["mail"][0].ToString()

            $count++
            Write-Progress -Activity ("Found mailbox (" + $count + "/" + $objects.Length + ")") -Status $mail -PercentComplete ($count/$objects.Length*100)

            $csv = ""
            $csv += '"' + $mail + '"' + ','		# Source Email
            $csv += '"",'						# Source UserName
            $csv += '"",'						# Source Password
            $csv += '"' + $mail + '"' + ','		# Destination Email
            $csv += '"",'						# Destination UserName
            $csv += '""'						# Destination Password

            Add-Content -Path $filename -Value $csv
        }
    }
}

######################################################################################################################################################
# Main menu -> Manage local Exchange Server -> Export public folder data
######################################################################################################################################################

function Action-ExportPublicFolderData([string]$typeOfProject)
{
    try
    {
        # Select an existing project
        Write-Host
        Write-Host -Object "Note: Please create an empty Public Folder project via MigrationWiz before using this option" -ForegroundColor Yellow
        $connector = Menu-MigrationWizConnectorListPrompt -connectors (MWHelper-GetConnectors)
        if($connector -eq $null)
        {
            Write-Host -Object "No project selected" -ForegroundColor Red
            return
        }
        if($connector.ExportConfiguration.ExchangeItemType -ne "PublicFolders")
        {	
            Write-Host -Object "Project '$($connector.Name)' is not a Public Folder project" -ForegroundColor Red
            return
        }
        $mailboxes = MW-GetMailboxes -connector $connector -mailboxOffSet 0 -mailboxPageSize 1
        if($mailboxes -ne $null)
        {
            Write-Host -Object "Project '$($connector.Name)' is not empty" -ForegroundColor Red
            return
        }

        # Connect to the exchange server
        switch($typeOfProject) {
            "Local" { ExchangeServerHelper-ConnectPowerShell }
            "O365" { Office365Helper-ConnectRemotePowerShell }
            default { Write-Host -Object "Type of project $typeOfProject is unknown" -ForegroundColor Red}
        }


        # Define the files that will hold the exported data
        $pfFileName = $env:temp + "\MigrationWiz-PublicFolders.json"
        $pfMailboxFileName = $env:temp + "\MigrationWiz-PublicFolderPerMailbox.json"
        $pfListCsvPath = $env:temp + "\MigrationWiz-PublicFolderList.csv"

        # TODO: Warn if overwriting existing files

        # Ask which folder to use as the root
        Write-Host -Object "Enter the root public folder being migrated below. Simply hit enter to use the top-most folder as the root (i.e. '\')"
        $rootIdentity = (Helper-PromptString -prompt "Root public folder path [\]" -allowEmpty $true)
        if($rootIdentity.Length -le 1)
        {
            $rootIdentity = "\"
        }

        # Retrieve all public folders
        Write-Host -Object "Retrieving all public folders ..."
        $publicFolders = (Get-PublicFolder -Identity $rootIdentity -ResultSize Unlimited -Recurse)
        Write-Host -Object "  Found $($publicFolders.Length) public folders ..."
        Write-Host -Object "Retrieving statistics on public folders ..."
        $publicFolderStatistics = (Get-PublicFolderStatistics -Identity $rootIdentity -ResultSize Unlimited)
        Write-Host -Object "  Found $($publicFolderStatistics.Length) statistics ..."

        # Store all public folders for processing
        Write-Host -Object "Outputting all public folders to $pfListCsvPath"
        $publicFolders | Select-Object Identity | Export-Csv -Encoding Unicode $pfListCsvPath
        MWHelper-CreateMailboxes -connector $connector -pfListCsvPath $pfListCsvPath
        
        # Change the Microsoft.Exchange.Data.Storage.Management.PublicFolder to BitTitan.ExchangeTools.PowerShell.Data.PublicFolderData
        $publicFolderDataCollection = New-Object -TypeName System.Collections.Generic.List[BitTitan.ExchangeTools.PowerShell.Data.PublicFolderData]
        foreach($publicFolder in $publicFolders)
        {
            $publicFolderData = New-Object -TypeName BitTitan.ExchangeTools.PowerShell.Data.PublicFolderData

            # Retrieve the statistics for the public folder
            $publicFolderStatistic = $publicFolderStatistics | Where-Object -Property EntryId -eq -Value $publicFolder.EntryId
            if ($publicFolderStatistic -ne $null)
            {
                $publicFolderData.FolderSize = [BitTitan.ItemSizeParser]::ParseItemSizeToBytes($publicFolderStatistic.TotalItemSize)
                $publicFolderData.ItemCount = $publicFolderStatistic.ItemCount
            }
            else
            {
                Write-Host "Cannot find public folder statistics for folder '$($publicFolder.Identity)'. Setting size to 0."
                $publicFolderData.FolderSize = $publicFolder.FolderSize
            }

            # Set the folder class/types
            if ($publicFolder.FolderType -ne $null)
            {
                $publicFolderData.FolderClass = $publicFolder.FolderType
            }
            else
            {
                $publicFolderData.FolderClass = $publicFolder.FolderClass
            }
            
            # Set the appropriate fields
            $publicFolderData.Identity = $publicFolder.Identity.ToString()
            $publicFolderData.ParentPath = $publicFolder.ParentPath
            $publicFolderData.IsValid = $publicFolder.IsValid
            $publicFolderData.MailEnabled = $publicFolder.MailEnabled
            $publicFolderData.Name = $publicFolder.Name

            # TODO: Retrieve the extra email addresses so that they can be set at the destination (see file history)
            
            # Add the data to the collection
            $publicFolderDataCollection.Add($publicFolderData)
        }

        # Write progress
        Write-Host -Object "Completed public folder lookup and object conversion ..."

        # Convert results to an array
        $publicFolderDataArray = $publicFolderDataCollection.ToArray()
        $publicFolderDataArray = $publicFolderDataArray | Sort-Object -Property Identity
        $arrayLength = $publicFolderDataArray.Length

        # Output the public folders to a JSON file
        Write-Host -Object "Outputting $arrayLength public folders to $pfFileName"
        ConvertTo-Json -InputObject $publicFolderDataArray -Compress -Depth 99 | Out-File -FilePath $pfFileName
        ExchangeServerHelper-SetPublicFolderFileName -fileName $pfFileName

        # Split the public folders into mailboxes
        Write-Host -Object "Splitting public folders into mailboxes"
        $publicFolderMailboxes = Split-PublicFoldersIntoMailboxes -MailboxSize 25 -PublicFolders $publicFolderDataArray

        # Output the public folder mailboxes
        Write-Host -Object "Outputting public folder mailboxes to $pfMailboxFileName"
        ConvertTo-Json -InputObject $publicFolderMailboxes -Compress -Depth 99 | Out-File -FilePath $pfMailboxFileName
        ExchangeServerHelper-SetPublicFolderPerMailboxFileName -fileName $pfMailboxFileName

        # Estimate the number of licenses
        Write-Host -Object "Estimating the number of Public Folder licenses required for the migration"
        $estimatedLicenses = Select-PublicFolderLicenseEstimate -PublicFolderData $publicFolderDataArray
        Write-Host -Object "License estimate: $estimatedLicenses public folder licenses"
    }
    finally
    {
        Get-PSSession | Remove-PSSession
    }
}

######################################################################################################################################################
# Main script execution
######################################################################################################################################################

try
{
    Clear-Host
    Helper-IncreaseWindowSize -width 120 -height 50
    Helper-LoadMigrationWizModule

    # Print banner
    Menu-Banner

    # Set environment
    MWHelper-ChooseEnvironment

    # Go to main loop
    if(MWHelper-GetTicket)
    {
        Menu-MainLoop
    }
}
catch
{
    throw
}