<# 
    .SYNOPSIS
    This script will add an admin user as owner to a given list of teams, or to all teams in the tenant.
    .EXAMPLE
    ./AddTeamsOwner.ps1 -username "admin@domain.com" -password "mypassword" -teamsIds ("id1", "id2")
    This adds user "admin@domain.com" to teams id1 and id2.
    .EXAMPLE
    ./AddTeamsOwner.ps1 -username "admin@domain.com" -password "mypassword" -all
    This adds user "admin@domain.com" to all teams in the tenant.
    .EXAMPLE
    get-help ./AddTeamsOwner.ps1
    This shows the help for the script.
#>

[CmdletBinding(PositionalBinding=$True)]
Param (
    # The admin username
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$username,

    # The admin password
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$password,

    # The team ids
    [Parameter(Mandatory=$false)]
    [string[]]$teamIds,

    # Whether to add owners to all teams
    [Parameter(Mandatory=$false)]
    [Switch]$all
)

$ErrorActionPreference = "Stop"

# Import module
Import-Module MicrosoftTeams

# Validate either $teamIds or $all is set
if (!$teamIds -and ($all -eq $false)) {
    throw "Please pass in a list of team ids to add the owner to, or use `-all` to add owners to all teams"
}

# Connect to MS Teams
$secPass = ConvertTo-SecureString $password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential ($username, $secPass)
Connect-MicrosoftTeams -Credential $credential 

# If a list of team ids are given, only add the admin as owner to these teams
if ($teamIds) {
    foreach ($teamId in $teamIds) {
        try {
            Remove-TeamUser -GroupId $teamId -User $username -Role Owner
        } catch {}
        Add-TeamUser -GroupId $teamId -User $username -Role Owner
        Write-Output "User $username was added as owner to team $teamId" 
    }
} 

# Else if $all is set, add owner to all teams
elseif ($all) {
    # Retrieve a list of teams that the admin is not an owner of
    $teamsToAddOwner = Get-Team | Where-Object { (Get-TeamUser -GroupId $_.GroupId -Role Owner | Where-Object { $_.User -eq $username }) -eq $null }

    # Add the admin as owner to these teams
    foreach ($team in $teamsToAddOwner) {
        Add-TeamUser -GroupId $team.GroupId -User $username -Role Owner
        Write-Output "User $username was added as owner to team $($team.GroupId)" 
    }
}
