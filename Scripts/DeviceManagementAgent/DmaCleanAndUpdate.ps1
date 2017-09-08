
<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			DmaCleanAndUpdateV4.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.03
	Date:			December 7, 2016

	Disclaimer: 		This script is provided ‘AS IS’. No warrantee is provided either expresses or implied.

	Copyright: 		Copyright© 2016 BitTitan. All rights reserved.
	
.Synopsis
   This script findsw and cleans up DMA Loggin Folders.

.NOTES
    This script must be run as an Administrator

.EXAMPLE
    .\DmaCleanAndUpdateV4.ps1 
#>

$ErrorActionPreference = "Stop"

$DmaService = Get-Service -Name "BitTitanDMA*"
if($null -eq $DmaService)
{
    Write-Error "The BitTitan Device Management Agent could not be found on this computer. The DMA must be installed before running this script"
}


# Confirm that the logged-in user has Administrator privileges. If not, throw error.
try
{
    $WindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $WindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal($WindowsIdentity)
    $AdministratorRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    if (-not $WindowsPrincipal.IsInRole($AdministratorRole))
    {
        Write-Error "This script must be run with Administrator privilieges. Right-Click the PowerShell Console and select 'Run as Administrator'."
    }
}
catch
{
    Write-Error "Could not obtain current User's Identity."
}




# Get the CPU architecture using IntPtr size, so we can be MUI-comliant (32-bit OR 64-bit)
$WmiArchitecture = [System.IntPtr]::Size
switch ($WmiArchitecture)
{
    4 
    {
        $bittitan_folder = Join-Path -Path $env:ProgramFiles -ChildPath "BitTitan"
    }


    8 
    {
        $bittitan_folder = Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath "BitTitan"
    }

    Default 
    {
        Write-Error "The CPU architecture is not supported."
    }
}

$dma_folder = Join-Path -Path $bittitan_folder -ChildPath "DeviceManagementAgent"
$log_folder = Join-Path -Path $dma_folder -ChildPath "log"
$updater_path = Join-Path -Path $dma_folder -ChildPath "BitTitanDMAUpdater.exe"
$updater_config_path = Join-Path -Path $dma_folder -ChildPath "updater.json"


# Stop the service
Write-Host "Stopping the DMA service...."
Stop-Service "BitTitanDMA" -Force


# Stop all the BittianDMA agents components and modules
Write-Host "Stopping DMA Processes..."
Get-Process "BitTitanDMA*" | Stop-Process -Force 


# Check and remove the log folder if it exists
Write-Host "Checking for the DMA log folders..."
if (Test-Path $log_folder -pathType container)
{
    try
    {
        Write-Host "Cleaning the logs..."
        Remove-Item $log_folder -Force -Recurse
        Write-Host "Logs were cleaned."
    }
    catch [Exception]
    {
        Write-Error "Failed to clean the logs - $($_.Exception.GetType().FullName) - $($_.Exception.Message)"
    }
}
else
{
    Write-Host "DMA log folder was not found. No need to clean."
}


Write-Host "Removing the updater config file..."
if (Test-Path $updater_config_path -pathType Leaf)
{
    Remove-Item $updater_config_path -Force
}


Write-Host "Launching the updater..."
if (Test-Path $updater_path -pathType Leaf)
{
    # Run the BitTitanDMAUpdater
    try
    {
        $updater_process = Start-Process -FilePath $updater_path -WorkingDirectory $dma_folder -PassThru -Wait
        Write-Host "The BitTitanDMAUpdater ran. Code:'$($updater_process.ExitCode)"
    }
    catch
    {
        Write-Error "Failed to run the BitTitanDMAUpdater. Exception: $($_.Exception.GetType().FullName) - $($_.Exception.Message)"
    }

    # Start the DMA Service
    try
    {
        Write-Host "Restarting the DMA service..."
        Start-Service "BitTitanDMA"
    }
    catch [Exception]
    {
        Write-Error "Failed to start the DMA service. Exception: $($_.Exception.GetType().FullName) - $($_.Exception.Message)"
    }

}
else 
{
    Write-Error "BitTitanDMAUpdater.exe was not found."
}