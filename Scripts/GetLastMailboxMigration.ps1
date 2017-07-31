<#
.NOTES
    Company:		BitTitan, Inc.
    Title:			GetLastMailboxMigration.PS1
    Author:			SUPPORT@BITTITAN.COM
    Requirements: 
    
    Version:		1.00
    Date:			April 19, 2017

    Exchange Version:	2016
    Windows Version:	WINDOWS 10 ENTERPRISE

    Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

    Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
    
.SYNOPSIS
    Retrieves the last mailbox migration entities for every mailbox.

.DESCRIPTION 	
    This script retrieves the last executed mailbox migration entity for every single mailbox of the given connector.

.INPUTS
    -[MigrationProxy.WebApi.Ticket] Ticket, the ticket for authentication.
    -[guid] ConnectorId, the id of the mailbox connector.
    -[string] Env, the context to work with. Valid options : BT, China.

.EXAMPLE
    .\GetLastMailboxMigration.ps1 -Ticket $MWTicket -ConnectorId '12345678-0000-0000-0000-000000000000' -Env 'BT' 
    Runs the script and outputs the last started mailbox migrations for each mailbox within connector 12345678-0000-0000-0000-000000000000.
#>

param(
    # Ticket for authentication
    [Parameter(Mandatory=$True)]
    [MigrationProxy.WebApi.Ticket] $Ticket,

    # The id of the connector
    [Parameter(Mandatory=$True)]
    [guid] $ConnectorId,
   
    # The environment to work with
    [Parameter(Mandatory=$False)]
    [ValidateSet("BT", "China")]
    [string] $Env = "BT"
) 

# Set page size
$PageSize = 100

# Build a container to save output
$outputDictionary = @{}

# Retrieve the mailbox migrations with pagination
Write-Verbose "Retrieving all mailbox migrations...This may take a while."
$count = 0
$mailboxMigrations = New-Object System.Collections.ArrayList
While($true)
{    
    [array]$temp = Get-MW_MailboxMigration -Ticket $Ticket -Environment $Env -FilterBy_Guid_ConnectorId $ConnectorId -PageOffset $($count*$PageSize) -PageSize $PageSize 
    $mailboxMigrations.AddRange($temp)
    if ($temp.count -lt $PageSize) { break } 
    $count++
}
Write-Verbose "Totally $($mailboxMigrations.Count) entities retrieved."

# Group the migrations by mailbox id and get the last one
foreach ($mailboxMigration in $mailboxMigrations)
{
    # If the mailbox id does not exist
    if (-not $outputDictionary.ContainsKey($mailboxMigration.MailboxId) -OR $outputDictionary[$mailboxMigration.MailboxId].StartDate -lt $mailboxMigration.StartDate)
    {
        $outputDictionary[$mailboxMigration.MailboxId] = $mailboxMigration
    }    
}

#Output
$outputDictionary.Values | Write-Output 