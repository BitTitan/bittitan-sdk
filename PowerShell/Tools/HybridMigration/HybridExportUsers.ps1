<#
.NOTES
    Company:          BitTitan, Inc.
    Title:            HybridExportUsers.ps1
    Date:             Sep 30, 2019
    Disclaimer:       This script is provided 'AS IS'. No warranty is provided either expressed or implied
    Copyright:        Copyright © 2019 BitTitan. All rights reserved
        
.SYNOPSIS
    Exports mailboxes information into a csv file for hybrid migration.

.DESCRIPTION
    Exports mailboxes information on the on-premises Exchange server into a csv file that can be used for users import for hybrid migration.

.INPUTS
    None.

.OUTPUTS
    MigrationWizHybridExportUsers.csv
        A csv file that contains the mailboxes information and can be used for user import for a hybrid migration.
    MigrationWizHybridExportUsers.log
        A log file that keeps track of the users export process.

#>

# Set the error action preference
$ErrorActionPreference = "Stop"

# Add the Exchange PowerShell snap-in to the current console
$exchangeSnapins = Get-PSSnapin -ErrorAction SilentlyContinue | Where-Object {$_.Name.Contains("Microsoft.Exchange.Management.PowerShell")}
if (-not $exchangeSnapins) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.*
}

# Create the log and output files
$logFile = "$(Get-Location)\migration_hybrid_export_users.log"
$outputFile = "$(Get-Location)\migration_hybrid_export_users.csv"
New-Item $logFile -ItemType file -force > $null
New-Item $outputFile -ItemType file -Force > $null
Write-Output "$outputFile has been created. It can be used for users import for hybrid migration."
Write-Output "$logFile has been created to keep track of the users export process."

# Get all mailboxes
$allMailboxes = Get-Mailbox -ResultSize unlimited
if (!$allMailboxes -or $mailboxCount -eq 0) {
    Write-Output "No mailboxes retrieved from the on-premises Exchange server." | Out-File -FilePath $logFile -Append -NoClobber
    exit
}

# Set the variables to keep track of the export process
$batchSize = 100
$batchProcess = 0
$lastIndex = $allMailboxes.Count - 1
$batchCount = [Math]::Ceiling($allMailboxes.Count / $batchSize)

# Loop to export all mailboxes in batch
while ($batchProcess -lt $batchCount) {
    # Get the index range of the mailboxes to be exported in the current batch
    $startIndex = $batchSize * $batchProcess
    $endIndex = $startIndex + $batchSize - 1
    if ($endIndex -gt $lastIndex) {
        $endIndex = $lastIndex
    }

    # Log the current batch
    Write-Output "Batch $($batchProcess)/$($batchCount): Performing export tasks for $($endIndex - $startIndex + 1) mailboxes." | Out-File -FilePath $logFile -Append -NoClobber

    # Get all the mailboxes to be exported in the current batch
    $mailboxBatch = $allMailboxes[$startIndex..$endIndex]

    # Create an empty list of objects
    $mailboxObjects = New-Object System.Collections.Generic.List[PSObject]

    # Loop to create an object with information needed for each mailbox
    foreach ($mailbox in $mailboxBatch) {
        # Create a new object
        $mailboxObject = New-Object PSObject

        # Add information to the object
        # Add the smtp address
        $mailboxObject | Add-Member NoteProperty -Name "SMTPAddress" -Value $mailbox.PrimarySmtpAddress

        # Get the organizatinal unit
        $organizationalUnit = $mailbox.OrganizationalUnit;

        # Remove the domain name in the organizational unit
        if ($organizationalUnit -and $organizationalUnit.IndexOf("/") -ge 0) {
            $organizationalUnit = [string]$organizationalUnit.Substring(($organizationalUnit.IndexOf("/") + 1))
        }

        # Add the organizatinal unit
        $mailboxObject | Add-Member NoteProperty -Name "OrganizationalUnit" -Value $organizationalUnit

        # Add the batch name and set it to be an empty string
        $mailboxObject | Add-Member NoteProperty -Name "BatchName" -Value ""

        # Get the cas mailbox properties
        $casMailbox = Get-CASMailbox -Identity $mailbox.PrimarySmtpAddress -ErrorAction SilentlyContinue
        
        # Add the imap, pop, owa and active sync
        if ($casMailbox) {
            $mailboxObject | Add-Member NoteProperty -Name "IMAP" -Value $([boolean]$casMailbox.ImapEnabled)
            $mailboxObject | Add-Member NoteProperty -Name "POP" -Value $([boolean]$casMailbox.PopEnabled)
            $mailboxObject | Add-Member NoteProperty -Name "OWA" -Value $([boolean]$casMailbox.OWAEnabled)
            $mailboxObject | Add-Member NoteProperty -Name "ActiveSync" -Value $([boolean]$casMailbox.ActiveSyncEnabled)
        }
        else {
            $mailboxObject | Add-Member NoteProperty -Name "IMAP" -Value ""
            $mailboxObject | Add-Member NoteProperty -Name "POP" -Value ""
            $mailboxObject | Add-Member NoteProperty -Name "OWA" -Value ""
            $mailboxObject | Add-Member NoteProperty -Name "ActiveSync" -Value ""

            # Log the error
            Write-Output "Failed to retrieve the cas mailbox properties for mailbox $($mailbox.PrimarySmtpAddress.Address)." | Out-File -FilePath $logFile -Append -NoClobber
        }
        
        # Add the object to the list of mailbox objects
        $mailboxObjects += $mailboxObject
    }

    # Convert the mailbox objects to csv
    $mailboxCsv = $mailboxObjects | ConvertTo-Csv -NoTypeInformation | ForEach-Object {$_ -replace '"',''}

    # Skip the header if it is not the first batch
    if ($batchProcess -gt 0) {
        $mailboxCsv = $mailboxCsv | Select-Object -Skip 1
    }

    # Save the mailbox csv to the output file
    $mailboxCsv | Out-File $outputFile -Append

    # Update the batch process
    $batchProcess++
}
