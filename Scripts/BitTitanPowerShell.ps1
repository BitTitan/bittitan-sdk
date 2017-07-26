<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			BitTitanPowerShell.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.00
	Date:			DECEMBER 1, 2016

	Windows Version:	WINDOWS 10 ENTERPRISE

	Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

	Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
	
.SYNOPSIS
	Imports the BitTitan powershell module.

.DESCRIPTION 	
	This script simply tries to import the BitTitanPowerShell.dll from the SDK installation folder.

.INPUTS	

.EXAMPLE
  	.\BitTitanPowerShell.ps1
	Imports the BitTitanPowerShell.dll into the context.
#>

$currentPath = Split-Path -parent $MyInvocation.MyCommand.Definition
Import-Module "$currentPath\BitTitanPowerShell.dll"

################################################################################
# Display MigrationWiz Commands Shortcut
################################################################################

function Get-MigrationWizCommands
{
	Get-Command -Module BitTitanPowerShell
}

################################################################################
# Increase Window Size
################################################################################

function Helper-IncreaseWindowSize([int]$width, [int]$height)
{
    # Returns if it is window size is null; this happens when running in PowerShell ISE
    if($host.ui.rawui.WindowSize -eq $null){
        return
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

################################################################################
# Display Instructions
################################################################################

Helper-IncreaseWindowSize 120 50

Write-Host
Write-Host -ForegroundColor White "+------------------------------------------------------------------------------+"
Write-Host -ForegroundColor White "| BitTitan Command Shell                                                       |"
Write-Host -ForegroundColor White "+------------------------------------------------------------------------------+"
Write-Host
Write-Host "Sample scripts can be found in the current folder.  Use these as building"
Write-Host "blocks to build your own."
Write-Host
Write-Host -NoNewline "Get help for a cmdlet     :"
Write-Host -NoNewline " "
Write-Host -ForegroundColor Yellow "Help <cmdlet name>"
Write-Host