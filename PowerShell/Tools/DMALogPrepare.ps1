<#
.NOTES
    Company:          BitTitan, Inc.
    Title:            DMALogPrepare.ps1
    Author:           Support@BitTitan.com

    Version:          1.01
    Date:             April 27, 2018

    Disclaimer:       This script is provided 'AS IS'. No warranty is provided either expressed or implied

    Copyright:        Copyright © 2018 BitTitan. All rights reserved

.SYNOPSIS
    Finds DMA Logs and places them in a zip file on the user's desktop

.DESCRIPTION
    This script will find the Device Management Agent Logs on the user's local machine, create a copy of them, strip 
    unnecessary files from the copy, and then zip the copy to the Desktop

#>

# Tests the log folder to be sure it exists
function CheckLogFolderExists($logFolder)
{
    return Test-Path -Path $logFolder
}

# Creates the zip archive and places it on the desktop
function ZipLogs($logFolder)
{
    Compress-Archive -Path $logFolder -DestinationPath "$logFolder.zip"
}

# Creates a directory to copy the logs to
function ProcessLogs($logFolder, $outputLogFolder)
{
    $excludedFileTypes = @("*.exe", "*.ps1", "*.bak", "*.dll")
    $folderToExclude = "BitTitan Powershell"
    New-Item -Path $outputLogFolder -ItemType directory | Out-Null
    Get-ChildItem $logFolder -Directory | Where-Object{$_.Name -notin $folderToExclude} | Copy-Item -Destination $outputLogFolder -Recurse -Exclude $excludedFileTypes -Force
}

# Deletes the folder containing the logs
function CleanUp($logFolder)
{
    Remove-Item -Path $logFolder -Recurse
}

# Main program
function Main
{
    try
    {
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

        # Add the logs to the zip archive
        ZipLogs $outputLogFolder

        # Delete the folder that the zip file was derived from
        CleanUp $outputLogFolder

        Write-Host "Please send the file, $logFolderName.zip, located on the Desktop, to BitTitan Support. Press RETURN to exit"
    }
    catch
    {
        Write-Host "Failed to get the log files."
        Write-Host $_.Exception.Message
        Write-Host "Press RETURN to exit."
    }
    finally
    {
        if (CheckLogFolderExists $outputLogFolder)
        {
            CleanUp $outputLogFolder
        }
        Read-Host
    }
}

Main