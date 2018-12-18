<#

.SYNOPSIS

88888888ba   88      888888888888  88                                            88                                
88      "8b  ""    ,d     88       ""    ,d                                      88                                
88      ,8P        88     88             88                                      88                                
88aaaaaa8P'  88  MM88MMM  88       88  MM88MMM  ,adPPYYba,  8b,dPPYba,           88  8b,dPPYba,    ,adPPYba,       
88""""""8b,  88    88     88       88    88     ""     `Y8  88P'   `"8a          88  88P'   `"8a  a8"     ""       
88      `8b  88    88     88       88    88     ,adPPPPP88  88       88          88  88       88  8b               
88      a8P  88    88,    88       88    88,    88,    ,88  88       88  "88     88  88       88  "8a,   ,aa  888  
88888888P"   88    "Y888  88       88    "Y888  `"8bbdP"Y8  88       88  d8'     88  88       88   `"Ybbd8"'  888  
                                                                        8"                                         
© Copyright 2018 BitTitan, Inc. All Rights Reserved.


.DESCRIPTION
    This script will start all the migrations created by Create-MW_AzureBlobContainerToOD4B.ps1 under a MSPC Customer.
    This script assumes that all migration items have corresponding subscriptions or licenses
	
.NOTES
	Author			For any questions contact Technical Sales Specialist Team <TSTeam@bittitan.com> or the author of this script Pablo Galan Sabugo <pablog@bittitan.com> 
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
	Change Log
        Please check the "Instructions" task of this Runbook for a detailed change log and full instructions
    Version: 1.1
#>


Function Get-CsvFile {
    Write-Host
    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the Azure Blob Container to Office 365 MigrationWiz project."
    Get-FileName $workingDir

    # Import CSV and validate if headers are according the requirements
    try {
        $lines = @(Import-Csv $global:inputFile)
    }
    catch {
        $msg = "ERROR: Failed to import '$global:inputFile' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit   
    }

    # Validate if CSV file is empty
    if ( $lines.count -eq 0 ) {
        $msg = "ERROR: '$global:inputFile' CSV file exist but it is empty. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Exit
    }

    # Validate CSV Headers
    $CSVHeaders = @("ProjectName")
    foreach ($header in $CSVHeaders) {
        if ($lines.$header -eq "" ) {
            $msg = "ERROR: '$global:inputFile' CSV file does not have all the required columns. Required columns are: '$($CSVHeaders -join "', '")'. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Exit
        }
    }

    Return $lines
 
 }

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################

# Import the functions
Import-Module .\BitTitanFunctions.psm1 -DisableNameChecking -force

#Working Directory
$global:workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Start-MW_AzureBlobContainerToO365Migrations.log"
$global:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($logFile)'."
Write-Host -ForegroundColor Yellow "         Generated CSV files will be in folder '$($workingDir)'."
Write-Host 
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg -LogFile $logFile

Connect-BitTitan

$connectorOffSet = 0
$connectorPageSize = 100
$connectors = $null

$connectorsFromCSVFile = @(Get-CsvFile)

Write-Host
$msg = "INFO: Retrieving Azure Blob Container to Office 365 MigrationWiz project..."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

if($connectorsFromCSVFile -ne $null -and $connectorsFromCSVFile.Length -ge 1) {
    Write-Host -ForegroundColor Green "SUCCESS: 1 Azure Blob Container to Office 365 MigrationWiz projects found:" 

        $connectorFromCSVFile = $connectorsFromCSVFile[0]
        Write-Host -Object "1-",$connectorFromCSVFile.ProjectName
}
else {
    $msg = "INFO: No Azure Blob Container to Office 365 MigrationWiz project found in the CSV file. Script aborted."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Exit
}

Write-Host

# Retrieve each connector
$connectors = @()
    try {
        $connectors += Get-MW_MailboxConnector -Ticket $mwTicket -Name $connectorFromCSVFile.ProjectName
    } catch {
        $msg = "ERROR: Failed to find the MigrationWiz project '$($connectorFromCSVFile.ProjectName)'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile  
        Exit
    }

$allMailboxes = @()
$mailboxes = @()
foreach ($connector in $connectors) {
    # Retrieve all mailboxes from the specified project
    $mailboxOffSet = 0
    $mailboxPageSize = 100
    $mailboxes = $null

    Write-Host "INFO: Retrieving migrations from '$($connector.Name)' MigrationWiz project"

    do {
        $mailboxesPage = @(Get-MW_Mailbox -Ticket $mwTicket -FilterBy_Guid_ConnectorId $connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize)

        if($mailboxesPage) {
            $mailboxes += @($mailboxesPage)
            $allMailboxes += @($mailboxesPage)

            foreach($mailbox in $mailboxesPage) {
                if($connector.ProjectType -eq "Archive") {
                    Write-Progress -Activity ("Retrieving migrations for '$($connector.Name)' Azure Blob Container to Office 365 MigrationWiz project") -Status $mailbox.ImportEmailAddress.ToLower()
                }
            }

            $mailboxOffSet += $mailboxPageSize
        }
    } while($mailboxesPage)

    if($mailboxes -ne $null -and $mailboxes.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($mailboxes.Length) migrations found." 
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No migrations found. Script aborted." 
        Exit
    }

}

# keep looping until specified to exit
do {
    $action = Menu-MigrationSubmission -AllMailboxes $allMailboxes -ProjectName "PST-O365-"
	if($action -ne $null) {
			$action = Menu-MigrationSubmission -AllMailboxes $allMailboxes -ProjectName "PST-O365-"
	}
	else {
	    Exit
	}
}
while($true)
