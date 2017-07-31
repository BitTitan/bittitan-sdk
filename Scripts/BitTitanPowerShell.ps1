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

################################################################################
# Load the BitTitan PowerShell Module
################################################################################

function Helper-LoadBitTitanModule()
{
    if (((Get-Module -Name "BitTitanPowerShell") -ne $null) -or ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null))
	{
		return;
	}

    $currentPath = Split-Path -parent $MyInvocation.MyCommand.Definition
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

################################################################################
# Display MigrationWiz Commands Shortcut
################################################################################

function Get-MigrationWizCommands
{
    if ((Get-InstalledModule -Name "BitTitanManagement" -ErrorAction SilentlyContinue) -ne $null) 
    {
        Get-Command -Module BitTitanManagement
    }
    else
    {
        Get-Command -Module BitTitanPowerShell
    }
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

Helper-LoadBitTitanModule
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