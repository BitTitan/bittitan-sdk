<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			SubmitFullMigration.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.00
	Date:			DECEMBER 1, 2016

	Exchange Version:	2016
	Windows Version:	WINDOWS 10 ENTERPRISE

	Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

	Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
	
.SYNOPSIS
	Initializes and starts a full mailbox migration process with the existing connector.

.DESCRIPTION 	
	This script will authenticate, retrieve the existing connector for the user and start the full migration process.

.INPUTS
	Inputs credential by the authentication dialog box (default).

.EXAMPLE
  	.\SubmitFullMigration.ps1
	Runs the script to start a full mailbox migration with the exiting connector.
#>

######################################################################################################################################
# Copyright © BitTitan 2016.  All rights reserved.
######################################################################################################################################

function SubmitFullMigration
{
	#import module 
	Import-MigrationWizModule

	#retrieve ticket
	$credentials = Get-Credential
	$ticket = Get-MW_Ticket -Credentials $credentials

	#retrieve connector 
	$connector =  Get-MW_MailboxConnector -Ticket $ticket -FilterBy_String_Name (Read-Host -Prompt "Connector")
	$items = Get-MW_Mailbox -ticket $ticket -FilterBy_Guid_ConnectorId $connector.Id 

	#start migration
	foreach ($item in $items)
	{
		Write-Host "Checking item" $item.ImportEmailAddress "with ID:" $item.Id 
		$result = Add-MW_MailboxMigration -Ticket $ticket -MailboxId $item.Id -Type Full -ConnectorId $connector.Id -UserId $ticket.UserId 
	}
}

######################################################################################################################################
# Helper functions.  																												 #
######################################################################################################################################

function Import-MigrationWizModule()
{
	if ((Get-Module -Name "BitTitanPowerShell") -ne $null)
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

SubmitFullMigration