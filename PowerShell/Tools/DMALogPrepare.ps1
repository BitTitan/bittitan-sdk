<#
.NOTES
    Company:          BitTitan, Inc.
    Title:            DMALogPrepare.ps1
    Author:           Support@BitTitan.com

    Version:          1.02
    Date:             June 25, 2018

    Disclaimer:       This script is provided 'AS IS'. No warranty is provided either expressed or implied

    Copyright:        Copyright © 2018 BitTitan. All rights reserved

.SYNOPSIS
    Finds DMA Logs and places them in a zip file on the user's desktop

.DESCRIPTION
    This script will find the Device Management Agent Logs on the user's local machine, create a copy of them, strip 
    unnecessary files from the copy, and then zip the copy to the Desktop
    Disclaimer: To bypass execution policy, run from batch with the command: "Powershell.exe -ExecutionPolicy Bypass -File .\DMALogPrepare.ps1"

#>

# Tests the log folder to be sure it exists
function CheckLogFolderExists($logFolder)
{
    return Test-Path -Path $logFolder
}

# This function looks for a compression assembly on the machine to use for zipping the folder.
function Find-CompressionAssembly
{
    try
    {
        Add-Type -AssemblyName "System.IO.Compression.FileSystem" -ErrorAction SilentlyContinue
        return $true
    }
    catch
    {
        try
        {
            Add-Type -Path "${env:ProgramFiles(x86)}\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\System.IO.Compression.FileSystem.dll" -ErrorAction SilentlyContinue
            return $true
        }
        catch
        {
            return $false
        }
    }
}

# Creates the zip archive and places it on the desktop
function ZipLogs($logFolder)
{
    # The Powershell version determines which features we can use.
    $psVersion = $PSVersionTable.PSVersion.Major

    # Version 2 or newer is required to run C# code
    if ($psVersion -lt 2)
    {
        Write-Host "Powershell version needs to be 2 or greater to run this script."
        return $false
    }

    # Versions 2 through 4 can run C# code
    elseif ($psVersion -le 4 -and $psVersion -ge 2)
    {
        # Attempt to load the Compression Assembly
        if (Find-CompressionAssembly -eq $false)
        {
            Write-Host "Could not zip the log files."
            return $false
        }

        try
        {
            [System.IO.Compression.ZipFile]::CreateFromDirectory($logFolder, "$logFolder.zip")
            return $true
        }
        catch
        {
            Write-Host "Could not zip the log files."
            return $false
        }
    }

    # Versions 5 and beyond have a newer commandlet
    else
    {
        try
        {
            Compress-Archive -Path $logFolder -DestinationPath "$logFolder.zip"
            return $true
        }
        catch
        {
            Write-Host "Could not zip the log files."
            return $false
        }
    }
}

# Creates a directory to copy the logs to
function ProcessLogs($logFolder, $outputLogFolder)
{
    $excludedFileTypes = @("*.exe", "*.ps1", "*.bak", "*.dll")
    New-Item -Path $outputLogFolder -ItemType directory | Out-Null

    Copy-Item -Path $logFolder -Destination $outputLogFolder -Recurse -Force

    Get-ChildItem -Path $outputLogFolder -Include $excludedFileTypes -Recurse | Remove-Item -Force
    
    if (CheckLogFolderExists "$outputLogFolder\BitTitan\BitTitan Powershell")
    {
        Remove-Item -Path "$outputLogFolder\BitTitan\BitTitan Powershell" -Recurse -Force
    }
}

# Deletes the folder containing the logs
function CleanUp($logFolder)
{
    Remove-Item -Path $logFolder -Recurse
}

# The manual steps to zip the folder, in case the automatic zipping fails.
function OutputManualSteps
{
    # Give the user the remaining steps, if the zip operation fails
	Write-Host "Failed to complete. Please perform the following steps:"
	Write-Host "1. Look for a folder named, $logFolderName, located on the Desktop."
	Write-Host '2. Right click on the folder. Hover over, "Send to" and choose "Compressed (zipped) folder".'
	Write-Host "3. There should now be a zipped folder named $logFolderName.zip on your desktop. Find it and send it to BitTitan Support."
	Write-Host "Press RETURN to exit."
}

# Main program
function Main
{
    try
    {
        Write-Host "Gathering the logs"

        $logTimeStamp = Get-Date -Format yyyy_MM_dd_tz_THH_mm_ss
        $logFolderPathProgramFiles = "${env:ProgramFiles(x86)}\BitTitan"
        $logFolderPathAppData = "${env:LOCALAPPDATA}\BitTitan"
        $logFolderName = "DMA Logs_$logTimeStamp"
        $outputLogFolder = "${env:USERPROFILE}\Desktop\$logFolderName"

        # Process the logs from Program Files x86 location, if it exists
        if (CheckLogFolderExists $logFolderPathProgramFiles)
        {
            ProcessLogs $logFolderPathProgramFiles "$outputLogFolder\ProgramFiles86"
        }
        else
        {
            Write-Host "No logs found within the Program Files (x86) folder"
        }
        
        # Process the logs from the Local App Data location, if it exists
        if (CheckLogFolderExists $logFolderPathAppData)
        {
            ProcessLogs $logFolderPathAppData "$outputLogFolder\LocalAppData"
        }
        else
        {
            Write-Host "No logs found within the LocalAppData folder"
        }

        Write-Host "Preparing the log files"
		
        # Add the logs to the zip archive
        $zipSuccess = ZipLogs $outputLogFolder
        if ($zipSuccess -eq $true)
        {
    	    # Delete the folder that the zip file was derived from
            CleanUp $outputLogFolder
   	        Write-Host "Please send the file, $logFolderName.zip, located on the Desktop, to BitTitan Support. Press RETURN to exit"
        }
        else
        {
            OutputManualSteps
        }
    }
    catch
    {
        Write-Host "Failed to get the log files. If this error persists, please notify BitTitan Support.`n"
        Write-Host $_.Exception.Message
        Write-Host "`n"
        OutputManualSteps
        Write-Host "Press RETURN to exit."
    }
    Read-Host
}

Main