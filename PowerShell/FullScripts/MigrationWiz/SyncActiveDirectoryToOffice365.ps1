<#
.NOTES
    Company:		BitTitan, Inc.
    Title:			SyncActiveDirectoryToOffice365.PS1
    Author:			SUPPORT@BITTITAN.COM
    Requirements: 
    
    Version:		1.00
    Date:			DECEMBER 1, 2016

    Exchange Version:	2016
    Windows Version:	WINDOWS 10 ENTERPRISE

    Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

    Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
    
.SYNOPSIS
    Synchronizes the Active Directory to Office 365.

.DESCRIPTION 	
    This script prompts the user to input credentials, select the environment, the types of object to synchronize and the types of synchronization. Then it processes the synchronization with logs.

.INPUTS	

.EXAMPLE
    .\SyncActiveDirectoryToOffice365.ps1
    Launches the script and helps users synchronize from AD to Office 365.
#>

# Active Directory domain controller to connect to.  Specify the FQDN.
# The default is that we will locate the closest domain controller.

$adServerName = ""

# Active Directory root search container for users.  The default is the root of the domain.
# i.e. OU=Users,DC=example,DC=com

$userRootSearchContainer = ""

# Upon user creation, a password needs to be set, you can specify a default password here.
# If no default password is specified, we will generate a strong random password for each
# user upon creation.

$userDefaultPassword = ""

# Active Directory root search container for contacts.  The default is the root of the domain.
# i.e. OU=Users,DC=example,DC=com

$contactRootSearchContainer = ""

# Active Directory root search container for groups.  The default is the root of the domain.
# i.e. OU=Users,DC=example,DC=com

$groupRootSearchContainer = ""

# User LDAP search filter.  Default value is:
# (&(objectCategory=person)(objectClass=user)(displayName=*)(mail=*)(userPrincipalName=*))

$userSearchFilter = ""

# Contact LDAP search filter.  Default value is:
# (&(objectCategory=person)(objectClass=contact)(displayName=*)(mail=*))

$contactSearchFilter = ""

# Group LDAP search filter.  Default value is:
# (&(objectClass=group)(displayName=*)(mail=*))

$groupSearchFilter = ""

# User exclusion filter.  This filter is applied to:
#
# 1. LDAP Path
# 2. Primary email address
# 3. User Principal Name
#
# Specify as a regular expression.  Default value is:
# (^DiscoverySearchMailbox|^SystemMailbox|^FederatedEmail)

$userExclusionFilter = ""

# Contact exclusion filter.  This filter is applied to:
#
# 1. LDAP Path
# 2. Primary email address
#
# Specify as a regular expression.  Default is no filter.

$contactExclusionFilter = ""

# Group exclusion filter.  This filter is applied to:
#
# 1. LDAP Path
# 2. Primary email address
#
# Specify as a regular expression.  Default is no filter.

$groupExclusionFilter = ""

#
# Do not modify any contents below this line.  You can make customizations
# to parameters listed above.
#

$migrationwizCreds = $null
$office365Creds = $null
$ticket = $null
$environment = $null

# Verbose and warning preferences are on by default
$VerbosePreference = 'Continue'
$WarningPreference = 'Continue'

# Debug and error preferences are off by default
$DebugPreference = 'SilentlyContinue'
$ErroractionPreference = 'SilentlyContinue'

function GetEnvironment()
{
    # Prompt for environment
    if($environment -eq $null)
    {
        $script:environment = Read-Host -Prompt "Select environment. Options include BT (default) or China. Press <Enter> to select default"
    }
    
    # Choose default environment
    if($environment -eq "")
    {
        $script:environment = "BT"
    }

    return $script:environment
}

function GetTicket()
{
    ################################################################################
    # prompt for credentials
    ################################################################################
    
    if($migrationwizCreds -eq $null)
    {
        # prompt for credentials
        $script:migrationwizCreds = $host.ui.PromptForCredential("BitTitan Credentials", "Enter your BitTitan user name and password", "", "")
    }

    if($migrationwizCreds)
    {
        # get new ticket if we don't already have one or it's expired
        if(($ticket -eq $null) -or ($ticket.ExpirationDate -lt (Get-Date).ToUniversalTime()))
        {
            # get new ticket
            $script:ticket = Get-MW_Ticket -Credentials $migrationwizCreds -Environment $script:environment
        }
    }

    return $ticket
}

function GetOffice365Credentials()
{
    ################################################################################
    # prompt for credentials
    ################################################################################

    if($office365Creds -eq $null)
    {
        # prompt for credentials
        $script:office365Creds = $host.ui.PromptForCredential("Office 365 Credentials", "Enter your Office 365 user name and password", "", "")
    }

    return $office365Creds
}

function GetObjectAction()
{
    Write-Host -Object "Select the object types to synchronize:" -ForegroundColor Yellow
    Write-Host
    Write-Host -Object "0 - Users, Contacts and Groups"
    Write-Host -Object "1 - Users only"
    Write-Host -Object "2 - Contacts only"
    Write-Host -Object "3 - Groups only"
    Write-Host -Object "x - Exit"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0 - 3"
        if($result -eq "x")
        {
            break
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 3))
        {
            return [int]$result
        }
    }

    return $null
}

function GetSyncAction()
{
    Write-Host -Object "Select the synchronization operation:" -ForegroundColor Yellow
    Write-Host
    Write-Host -Object "0 - Simulate without delete"
    Write-Host -Object "1 - Simulate with delete"
    Write-Host -Object "2 - Sync without delete"
    Write-Host -Object "3 - Sync with delete"
    Write-Host -Object "x - Exit"
    Write-Host

    while($true)
    {
        $result = Read-Host -Prompt "Select 0 - 3"
        if($result -eq "x")
        {
            break
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 3))
        {
            return [int]$result
        }
    }

    return $null
}

function SyncAll([MigrationProxy.WebApi.Ticket]$ticket, $simulate, $delete)
{
    SyncUsers -ticket $ticket -simulate $simulate -delete $delete
    SyncContacts -ticket $ticket -simulate $simulate -delete $delete
    SyncGroups -ticket $ticket -simulate $simulate -delete $delete
}

function SyncUsers([MigrationProxy.WebApi.Ticket]$ticket, $simulate, $delete)
{
    New-ActiveDirectoryToOffice365UserDirectorySync -Credentials $migrationwizCreds -Office365Credentials $office365Creds -EnableDeletes $delete -SimulationOnly $simulate -ServerName $adServerName -RootSearchContainer $userRootSearchContainer -SearchFilter $userSearchFilter -ExclusionFilter $userExclusionFilter -DefaultPassword $userDefaultPassword -Environment (GetEnvironment)
}

function SyncContacts([MigrationProxy.WebApi.Ticket]$ticket, $simulate, $delete)
{
    New-ActiveDirectoryToOffice365ContactDirectorySync -Credentials $migrationwizCreds -Office365Credentials $office365Creds -EnableDeletes $delete -SimulationOnly $simulate -ServerName $adServerName -RootSearchContainer $contactRootSearchContainer -SearchFilter $contactSearchFilter -ExclusionFilter $contactExclusionFilter -Environment (GetEnvironment)
}

function SyncGroups([MigrationProxy.WebApi.Ticket]$ticket, $simulate, $delete)
{
    New-ActiveDirectoryToOffice365GroupDirectorySync -Credentials $migrationwizCreds -Office365Credentials $office365Creds -EnableDeletes $delete -SimulationOnly $simulate -ServerName $adServerName -RootSearchContainer $groupRootSearchContainer -SearchFilter $groupSearchFilter -ExclusionFilter $groupExclusionFilter -Environment (GetEnvironment)
}

&{
    Clear-Host
    Write-Host
    
    # Get the environment for credentials
    GetEnvironment

    if(GetTicket)
    {
        if(GetOffice365Credentials)
        {
            # keep looping until specified to exit object action
            do
            {
                Write-Host
                $objectAction = GetObjectAction
                if($objectAction -ne $null)
                {
                    # keep looping until specified to exit sync action
                    do
                    {
                        Write-Host
                        $syncAction = GetSyncAction
                        if($syncAction -ne $null)
                        {
                            Write-Host

                            $simulate = $true
                            $delete = $false

                            switch($syncAction)
                            {
                                0 # Simulate without delete
                                {
                                    $simulate = $true
                                    $delete = $false
                                }

                                1 # Simulate with delete
                                {
                                    $simulate = $true
                                    $delete = $true
                                }

                                2 # Sync without delete
                                {
                                    $simulate = $false
                                    $delete = $false
                                }

                                3 # Sync with delete
                                {
                                    $simulate = $false
                                    $delete = $true
                                }
                            }

                            # Start transcript on specified log file
                            $logFilename = "Sync-ActiveDirectoryToOffice365." + (Get-Date).ToString("yyyyMMdd.HHmmss") + ".log"
                            Start-Transcript -path $logFilename -Append

                            switch($objectAction)
                            {
                                0 # All objects
                                {
                                    SyncAll -ticket (GetTicket) -simulate $simulate -delete $delete
                                }

                                1 # Users only
                                {
                                    SyncUsers -ticket (GetTicket) -simulate $simulate -delete $delete
                                }

                                2 # Contacts only
                                {
                                    SyncContacts -ticket (GetTicket) -simulate $simulate -delete $delete
                                }

                                3 # Groups only
                                {
                                    SyncGroups -ticket (GetTicket) -simulate $simulate -delete $delete
                                }
                            }
                            Stop-Transcript
                        }
                    }
                    while($syncAction -ne $null)
                }
            }
            while($objectAction -ne $null)
        }
    }
}
trap
{
    break;
}

Write-Host