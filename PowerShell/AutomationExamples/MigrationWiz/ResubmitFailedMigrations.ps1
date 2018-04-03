<#
.NOTES
    Company:        BitTitan, Inc.
    Title:          ResubmitFailedMigrations.PS1
    Author:         SUPPORT@BITTITAN.COM
    Requirements: 
    
    Version:        1.00
    Date:           April 2, 2018

    Disclaimer:     This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

    Copyright:      Copyright © 2018 BitTitan. All rights reserved.
    
.SYNOPSIS
    Resubmits all the migrations in a project that are in the failed status.

.DESCRIPTION 	
    This script will authenticate, retrieve all the mailboxes in a project and resubmit them if their last status is failed.

.INPUTS
    Inputs credential by the authentication dialog box (default).

.EXAMPLE
    .\ResubmitFailedMigrations.ps1
    Runs the script to resubmit all the migrations in a failed status.
#>

######################################################################################################################################
# Copyright © BitTitan 2018.  All rights reserved.
######################################################################################################################################
function ResubmitFailedMigrations
{
    # Import module 
    Import-MigrationWizModule

    # Retrieve ticket
    $credentials = Get-Credential
    $ticket = Get-MW_Ticket -Credentials $credentials

    # Retrieve connector
    $connector = Get-MW_MailboxConnector -Ticket $ticket -PageSize 1 -Name (Read-Host -Prompt "Project name")

    # Retrieve all items
    $items = Get-MW_Mailbox -Ticket $ticket -RetrieveAll -ConnectorId $connector.Id 

    # Loop through all items
    foreach ($item in $items)
    {
        Write-Host "Checking item" $item.ImportEmailAddress "with ID:" $item.Id

        # Retrieve status of the last submission
        $lastMigrationAttempt = Get-MW_MailboxMigration -Ticket $ticket -MailboxId $item.Id -PageSize 1 -SortBy_CreateDate_Descending
        
        # Check if last submission failed
        if ($lastMigrationAttempt.Status -eq "Failed")
        {
            Write-Host "Resubmitting Item" $item.ImportEmailAddress  -foregroundcolor red -backgroundcolor green
            
            # Resubmit the migrations with same parameters as the previous submission
            $result = Add-MW_MailboxMigration -ticket $ticket -MailboxId $lastMigrationAttempt.MailboxId -Type $lastMigrationAttempt.Type -ConnectorId $connector.Id -UserId $ticket.UserId -Status Submitted -ItemTypes $lastMigrationAttempt.ItemTypes
        }
    }
}

######################################################################################################################################
# Import the BitTitanPowerShell in the current session                                                                               #
######################################################################################################################################
function Import-MigrationWizModule()
{
    # Check if the BitTitanPowerShell module is already loaded in the current session or installed
    if (((Get-Module -Name "BitTitanPowerShell") -ne $null) -or ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null))
    {
        return;
    }

    # Build a search path
    $currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
    $moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")

    # Loop through all possible locations
    foreach ($moduleLocation in $moduleLocations)
    {
        # Check if folder exists
        if (Test-Path $moduleLocation)
        {
            # Import the module
            Import-Module -Name $moduleLocation
            return
        }
    }
    Write-Error "BitTitanPowerShell module was not loaded"
}

# Call main function
ResubmitFailedMigrations
