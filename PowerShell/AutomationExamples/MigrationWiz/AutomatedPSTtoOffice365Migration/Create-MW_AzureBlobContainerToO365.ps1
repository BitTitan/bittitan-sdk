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
    This script will create a MigrationWiz project to migrate all PST files uploaded by UploaderWiz to the corresponding
    Office 365 mailboxes. It will generate a CSV file with the MigrationWiz project name that will be used by the script 
    Start-MW_AzureBlobContainerToOD4BMigrations.ps1 to submit all the migrations.
	
.NOTES
	Author			For any questions contact Technical Sales Specialist Team <TSTeam@bittitan.com> or the author of this script Pablo Galan Sabugo <pablog@bittitan.com> 
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    BitTitan cannot be held responsible for any misuse of the script.
	Change Log
        Please check the "Instructions" task of this Runbook for a detailed change log and full instructions
    Version: 1.1
#>


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
$logFileName = "$(Get-Date -Format yyyyMMdd)_Create-MW_FileServerToOD4B.log"
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
$exportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "source" -EndpointType "Pst"
#Get source endpoint credentials
[PSObject]$exportEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $exportEndpointId 

#Select destination endpoint
$importEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "destination" -EndpointType "ExchangeOnline2"
#Get source endpoint credentials
[PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $importEndpointId 

#Create AzureFileSystem-OneDriveProAPI Document project
Write-Host
$msg = "INFO: Creating MigrationWiz PST to Office 365 project."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

$ProjectName = "PST-O365-$(Get-Date -Format yyyyMMddHHmm)"

# Export data
if(!$exportEndpointData.ContainerName){
    $containerName = "migrationwizpst"
}
else {
    $containerName = $exportEndpointData.ContainerName
}

$exportType = "Pst"
$exportTypeName = "MigrationProxy.WebApi.AzureConfiguration"
$exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
    "AdministrativeUsername" = $exportEndpointData.AdministrativeUsername;
    "AccessKey" = $exportEndpointData.AccessKey;
    "ContainerName" = $containerName;
    "UseAdministrativeCredentials" = $true
}
$exportEndpointId = $exportEndpointId
# Import data
$importType = "ExchangeOnline2"
$importTypeName = "MigrationProxy.WebApi.ExchangeConfiguration"
$importConfiguration = New-Object -TypeName $importTypeName -Property @{
    "Url" = $importEndpointData.Url;
    "AdministrativeUsername" = $importEndpointData.AdministrativeUsername;
    "AdministrativePassword" = $importEndpointData.AdministrativePassword;
    "UseAdministrativeCredentials" = $true
}
$importEndpointId = $importEndpointId

$ProjectType = "Archive"
$maximumSimultaneousMigrations = 400
$advancedOptions = "UseEwsImportImpersonation=1"

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
-maximumSimultaneousMigrations $maximumSimultaneousMigrations

Write-Host
$msg = "ACTION: Click on 'Autodiscover Items' directly in MigrationWiz to import the PST files into the MigrationWiz project."
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message $msg -LogFile $logFile

$url = "https://migrationwiz.bittitan.com/app/projects/$connectorId`?qp_currentWorkgroupId=$workgroupId"

$msg = "INFO: Opening '$url' in your default web browser."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

$result= Start-Process $url
Start-Sleep 5
WaitForKeyPress -Message "ACTION: If you have imported the PST files into the MigrationWiz project '$ProjectName', press any key to continue"
Write-Host


$msg = "ACTION: Apply User Migration Bundle licenses to the Office 365 email addresses in MSPComplete."
Write-Host -ForegroundColor Yellow $msg
Log-Write -Message $msg -LogFile $logFile

$customerUrlId = Get-CustomerUrlId -CustomerOrganizationId $customerOrganizationId

$url = "https://manage.bittitan.com/customers/$customerUrlId/users?qp_currentWorkgroupId=$workgroupId"

$msg = "INFO: Opening '$url' in your default web browser."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

$result= Start-Process $url
Start-Sleep 5
WaitForKeyPress -Message "ACTION: If you have applied the User Migration Bundle to the users, press any key to continue"
Write-Host


try {
    @($ProjectName) | Select-Object @{Name='ProjectName';Expression={$_}} | Export-Csv -Path $workingDir\PSTtoO365Project.csv -NoTypeInformation -force

    $msg = "SUCCESS: CSV file with the script output '$workingDir\PSTtoO365Project.csv' opened."
    Write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg -LogFile $logFile

    #Open the CSV file
    Start-Process -FilePath $workingDir\PSTtoO365Project.csv

    $msg = "INFO: This CSV file will be used by Start-MW_ AzureBlobContainerToO365Migrations.ps1 script to automatically submit all PST files for migration."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host
}
catch {
    $msg = "ERROR: Failed to export and open '$workingDir\PSTtoO365Project.csv' CSV file."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg -LogFile $logFile

##END SCRIPT