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
    This script will create a MigrationWiz project to migrate FileServer Home Directories to DropBox accounts.
    It will generate a CSV file with the MigrationWiz project and all the migrations that will be used by the script 
    Start-MW_FileServerToDropBox.ps1 to submit all the migrations.
	
.NOTES
	Author			For any questions contact Technical Sales Specialist Team <TSTeam@bittitan.com> or the author of this script Pablo Galan Sabugo <pablog@bittitan.com> 
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
#>


Function Get-CsvFile {
    Write-Host
    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the Home Directories and Dropbox email addresses."
    Get-FileName -InitialDirectory $PSScriptRoot

    # Import CSV and validate if headers are according the requirements
    try {
        $lines = @(Import-Csv $global:inputFile)
    }
    catch {
        $msg = "ERROR: Failed to import '$script:inputFile' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit   
    }

    # Validate if CSV file is empty
    if ( $lines.count -eq 0 ) {
        $msg = "ERROR: '$script:inputFile' CSV file exist but it is empty. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Exit
    }

    # Validate CSV Headers
    $CSVHeaders = "SourceFolder,DropBoxEmailAddress"
    foreach ($header in $CSVHeaders) {
        if ($lines.$header -eq "" ) {
            $msg = "ERROR: '$script:inputFile' CSV file does not have all the required columns. Required columns are: '$($CSVHeaders -join "', '")'. Script aborted."
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
$logFileName = "$(Get-Date -Format yyyyMMdd)_Create-MW_AzureBlobContainerToDropBox.log"
$global:logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($logFile)'."
Write-Host -ForegroundColor Yellow "         Generated CSV file will be in folder '$($workingDir)'."
Write-Host 
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg -LogFile $logFile 

Connect-BitTitan

#Select workgroup
$WorkgroupId = Select-MSPC_WorkGroup 

#Select customer
$customerOrganizationId = Select-MSPC_Customer -Workgroup $WorkgroupId

#Select source endpoint
$exportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "source" -EndpointType "AzureFileSystem"
#Get source endpoint credentials
[PSObject]$exportEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $exportEndpointId 

#Select destination endpoint
$importEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "destination" -EndpointType "DropBox"
#Get source endpoint credentials
[PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $importEndpointId 

$users = @(Get-CsvFile)
$totalLines = $users.count
 
#Create AzureFileSystem-Dropbox Document project
Write-Host
$msg = "INFO: Creating MigrationWiz FileServer to DropBox project."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

$ProjectName = "FS-DropBox-$(Get-Date -Format yyyyMMddHHmm)"
$ProjectType = "Storage"   
$exportType = "AzureFileSystem" 
$importType = "DropBox"

$exportEndpointId = $exportEndpointId
$importEndpointId = $importEndpointId

$exportTypeName = "MigrationProxy.WebApi.AzureConfiguration"
$exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
    "AdministrativeUsername" = $exportEndpointData.AdministrativeUsername;
    "AccessKey" = $exportEndpointData.AccessKey;
    "UseAdministrativeCredentials" = $true
}

$importTypeName = "MigrationProxy.WebApi.DropBoxConfigration"
$importConfiguration = New-MW_DropBoxConfiguration -UseAdministrativeCredentials $true -AdministrativePassword ""

$advancedOptions = "InitializationTimeout=28800000 FolderLimit=20000"

$connectorId = Create-MW_Connector -CustomerOrganizationId $customerOrganizationId `
-ProjectName $ProjectName `
-ProjectType $ProjectType `
-importType $importType `
-exportType $exportType `
-exportEndpointId $exportEndpointId `
-importEndpointId $importEndpointId `
-exportConfiguration $exportConfiguration `
-importConfiguration $importConfiguration `
-advancedOptions $advancedOptions `
-maximumSimultaneousMigrations 100

$msg = "INFO: Adding advanced options '$advancedOptions' to the project."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile 

$applyCustomFolderMapping = $false
do {
    $confirm = (Read-Host -prompt "Do you want to add a custom folder mapping to move the home directory under a folder?  [Y]es or [N]o")

    if($confirm.ToLower() -eq "y") {
        $applyCustomFolderMapping = $true
        
        do {
            Write-host -ForegroundColor Yellow  "ACTION: Enter the destination folder name: "  -NoNewline
            $destinationFolder = Read-Host

        } while($destinationFolder -eq "")
        
    }

} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

$msg = "INFO: Adding migrations to the project:"
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

$processedLines = 0
$FileServerDropBoxProject = @()

foreach ($user in $users) {        
    $SourceFolder= $user.SourceFolder
    $importEmailAddress =  $user.DropBoxEmailAddress 

    if($SourceFolder -ne "" -and $importEmailAddress -ne "") {

        #Double Quotation Marks
        [string]$CH34=[CHAR]34
        $folderFilter = "^((?!" + $SourceFolder + ").)*$"
        if ($applyCustomFolderMapping) {
            $folderMapping= "FolderMapping=" + $CH34 + "^" + $user.SourceFolder + "->" + $destinationFolder + $CH34
        }else {
            $folderMapping= "FolderMapping=" + $CH34 + "^" + $user.SourceFolder + "->" + $CH34
        }

        try {
            $result = Add-MW_Mailbox -ticket $global:mwTicket -ConnectorId $connectorId  -ImportEmailAddress $importEmailAddress -FolderFilter $folderFilter –AdvancedOptions $folderMapping

            $tab = [char]9
            $msg = "SUCCESS: Migration '$SourceFolder->$importEmailAddress'$tab with folder filter '$folderFilter' and folder mapping '$folderMapping'."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg -LogFile $logFile

            $ProcessedLines += 1

            [array]$FileServerDropBoxProject += New-Object PSObject -Property @{ProjectName=$ProjectName;SourceFolder=$SourceFolder;EmailAddress=$importEmailAddress} 
        }
        catch {
            $msg = "ERROR: Failed to add source folder and destination primary SMTP address." 
            write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg -LogFile $logFile  
            Exit
        }

    }
    else{
        if($SourceFolder -eq "") {
            $msg = "ERROR: Missing source folder in the CSV file. Skipping '$importEmailAddress' user processing."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Continue    
        } 
        if($importEmailAddress -eq "") {
            $msg = "ERROR: Missing destination OneDrive For Business email address in the CSV file. Skipping '$SourceFolder' source folder processing."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Continue    
        }         
    }              
}

write-Host
$msg = "SUCCESS: $ProcessedLines out of $totalLines source folders have been processed and added to the project '$ProjectName'." 
write-Host -ForegroundColor Green $msg
Log-Write -Message $msg -LogFile $logFile

Write-Host
$msg = "ACTION: Click on 'Request Access Token' directly in MigrationWiz to complete the DropBox authorization."
Write-Host -ForegroundColor Yellow $msg

$url = "https://migrationwiz.bittitan.com/app/projects/$connectorId/edit?qp_currentWorkgroupId=$workgroupId&stepKey=projectAuthorization"

Log-Write -Message $msg -LogFile $logFile
$msg = "INFO: Opening '$url' in your default web browser."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

$result= Start-Process $url
Start-Sleep 5
WaitForKeyPress -Message "ACTION: If you have clicked on 'Request Access Token' and a Success status is returned, press any key to continue"

Write-Host

$customerUrlId = Get-CustomerUrlId -CustomerOrganizationId $customerOrganizationId

$url = "https://manage.bittitan.com/customers/$customerUrlId/users?qp_currentWorkgroupId=$workgroupId"
$msg = "ACTION: Apply User Migration Bundle licenses to the OneDrive For Business email addresses in MSPComplete."
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message $msg -LogFile $logFile
$msg = "INFO: Opening '$url' in your default web browser."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

$result= Start-Process $url
Start-Sleep 5
WaitForKeyPress -Message "ACTION: If you have applied the User Migration Bundle to the users, press any key to continue"
Write-Host

try {
    $FileServerDropBoxProject| Select-Object ProjectName,SourceFolder,EmailAddress | sort { $_.EmailAddress } |Export-Csv -Path "$workingDir\FileServerDropBoxProject.csv" -NoTypeInformation -force
    
    $msg = "SUCCESS: CSV file with the script output '$workingDir\FileServerDropBoxProject.csv' opened."
    Write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg -LogFile $logFile

    #Open the CSV file
    Start-Process -FilePath "$workingDir\FileServerDropBoxProject.csv"

    $msg = "INFO: This CSV file will be used by Start-MW_FileServerToDropBoxMigrations.ps1 script to automatically submit all home directories for migration."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host
}
catch {
    $msg = "ERROR: Failed to export and open '$workingDir\FileServerDropBoxProject.csv' CSV file."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg -LogFile $logFile

##END SCRIPT