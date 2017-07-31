<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			PauseMigration.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.00
	Date:			DECEMBER 1, 2016

	Windows Version:	WINDOWS 10 ENTERPRISE

	Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

	Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
	
.SYNOPSIS
	Puts the last mailbox migration process in pause.

.DESCRIPTION 	
	This script will authenticate, retrieve the last mailbox migration process for the user and set the migration process status to pause.

.INPUTS
	Inputs credential by the authentication dialog box (default).

.EXAMPLE
  	.\PauseMigration.ps1
	Runs the script to pause the last mailbox migration process.
#>

######################################################################################################################################
# Copyright © BitTitan 2016.  All rights reserved.
######################################################################################################################################

function PauseMigration
{
	# Import module 
	Import-MigrationWizModule

	#retrieve ticket
	$credentials = Get-Credential
	$ticket = Get-MW_Ticket -Credentials $credentials

	#retrieve connector
	$connector = (Get-MW_MailboxConnector -Ticket $ticket -FilterBy_String_Name (Read-Host -Prompt "Connector")) | Select -First 1

	#retrieve items
	$items = Get-MW_Mailbox -Ticket $ticket -FilterBy_Guid_ConnectorId $connector.Id 

	#pause items
	foreach ($item in $items)
	{
		Write-Host "Pausing item" $item.ImportEmailAddress "with ID:" $item.Id 
		$lastMigrationAttempt = (Get-MW_MailboxMigration -Ticket $ticket -FilterBy_Guid_MailboxId $item.Id -SortBy_CreateDate_Descending) | Select -First 1
		if ($lastMigrationAttempt.Status -eq "Processing")
		{
			Set-MW_MailboxMigration -Ticket $ticket -mailboxmigration $lastMigrationAttempt -Status Stopping    
		}
	}
}

######################################################################################################################################
# Helper functions.  																												 #
######################################################################################################################################

function Import-MigrationWizModule()
{
	if (((Get-Module -Name "BitTitanPowerShell") -ne $null) -or ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null))
	{
		return;
	}

	$currentPath = Split-Path -parent $script:MyInvocation.MyCommand.Definition
	$moduleLocations = @("$currentPath\BitTitanPowerShell.dll", "$env:ProgramFiles\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll",  "${env:ProgramFiles(x86)}\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll")
	foreach ($moduleLocation in $moduleLocations)
	{
		if (Test-Path $moduleLocation)
		{
			Import-Module -Name $moduleLocation
			return
		}
	}
	
	Write-Error "BitTitanPowerShell module was not loaded"
}

PauseMigration