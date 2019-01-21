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
    This script will create a MigrationWiz project to migrate FileServer Home Directories to OneDrive For Business accounts.
    It will generate a CSV file with the MigrationWiz project and all the migrations that will be used by the script 
    Start-MW_FileServerToOD4B.ps1 to submit all the migrations.
	
.NOTES
	Author			For any questions contact Technical Sales Specialist Team <TSTeam@bittitan.com> or the author of this script Pablo Galan Sabugo <pablog@bittitan.com> 
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    BitTitan cannot be held responsible for any misuse of the script.
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
$logFileName = "$(Get-Date -Format yyyyMMdd)_Create-MW_AzureBlobContainerToOD4B.log"
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

write-host

$od4bv2 = $false
do {
    $confirm = (Read-Host -prompt "Do you want to use the new OneDrive For Business v2 endpoint (it requires an Azure subscription)?  [Y]es or [N]o")

    if($confirm.ToLower() -eq "y") {
        $od4bv2 = $true    

        #Select destination endpoint
        $importEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "destination" -EndpointType "OneDriveProAPI"
        #Get source endpoint credentials
        [PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $importEndpointId 
    }
    else {
        #Select destination endpoint
        $importEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "destination" -EndpointType "OneDrivePro"
        #Get source endpoint credentials
        [PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $importEndpointId 
    }

} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

#Create a PSCredential object to connect to Azure Active Directory tenant
$administrativeUsername = $importEndpointData.AdministrativeUsername
$administrativePassword = ConvertTo-SecureString -String $($importEndpointData.AdministrativePassword) -AsPlainText -Force
$o365Credentials = New-Object System.Management.Automation.PSCredential ($administrativeUsername, $administrativePassword)
write-host
$od4bArray = Get-OD4BAccounts -Credentials $o365Credentials

#Export users with OD4B to CSV file
$od4bArray | Select-Object SourceFolder,UserPrincipalName | sort { $_.UserPrincipalName } | Export-Csv -Path $workingDir\OD4BAccounts.csv -Delimiter "," -Encoding UTF8 -NoTypeInformation -Force #-Append

#Open the CSV file
try {
    
    Start-Process -FilePath $workingDir\OD4BAccounts.csv

    $msg = "SUCCESS: CSV file '$workingDir\OD4BAccounts.csv' processed, exported and open."
    Write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg -LogFile $logFile
}
catch {
    $msg = "ERROR: Failed to open '$workingDir\OD4BAccounts.csv' CSV file."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}

WaitForKeyPress -Message "ACTION: If you have edited and saved the CSV file then press any key to continue." 

#Re-import the edited CSV file
Try{
    $users = @(Import-CSV "$workingDir\OD4BAccounts.csv" | where-Object { $_.PSObject.Properties.Value -ne ""} | sort { $_.UserPrincipalName.length } )
    $totalLines = $users.Count

    if($users -eq $null) {
        $msg = "INFO: No Office 365 users found with OneDrive For Business. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Exit
    } 
}
Catch [Exception] {
    $msg = "ERROR: Failed to import the CSV file '$workingDir\OD4BAccounts.csv'."
    Write-Host -ForegroundColor Red  $msg
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $msg -LogFile $logFile
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}

#Create AzureFileSystem-OneDriveProAPI Document project
Write-Host
$msg = "INFO: Creating MigrationWiz FileServer to OneDrive For Business project."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

$destinationDomain = @((Get-MsolDomain).Name)[0]
$ProjectName = "FS-OD4B-$destinationDomain-$(Get-Date -Format yyyyMMddHHmm)"
$ProjectType = "Storage"   
$exportType = "AzureFileSystem" 
if($od4bv2) {
    $importType = "OneDriveProAPI"
}
else {
    $importType = "OneDrivePro"
}

$exportEndpointId = $exportEndpointId
$importEndpointId = $importEndpointId

$exportTypeName = "MigrationProxy.WebApi.AzureConfiguration"
$exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
    "AdministrativeUsername" = $exportEndpointData.AdministrativeUsername;
    "AccessKey" = $exportEndpointData.AccessKey;
    "UseAdministrativeCredentials" = $true
}

if($od4bv2) {
    $importTypeName = "MigrationProxy.WebApi.SharePointOnlineConfiguration"
    $importConfiguration = New-Object -TypeName $importTypeName -Property @{
        "AdministrativeUsername" = $importEndpointData.AdministrativeUsername;
        "AdministrativePassword" = $importEndpointData.AdministrativePassword;
        "AzureAccountKey" = $importEndpointData.AzureAccountKey;
        "AzureStorageAccountName" = $importEndpointData.AzureStorageAccountName;
        "UseAdministrativeCredentials" = $true
    }
}
else{
    $importTypeName = "MigrationProxy.WebApi.SharePointOnlineConfiguration"
    $importConfiguration = New-Object -TypeName $importTypeName -Property @{
        "AdministrativeUsername" = $importEndpointData.AdministrativeUsername;
        "AdministrativePassword" = $importEndpointData.AdministrativePassword;
        "UseAdministrativeCredentials" = $true
    }
}

#$advancedOptions = "InitializationTimeout=8 RenameConflictingFiles=1 ShrinkFoldersMaxLength=200"

$advancedOptions = "InitializationTimeout=8 RenameConflictingFiles=1 IncreasePathLengthLimit=1 SyncItems=1"

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
-maximumSimultaneousMigrations $totalLines

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
$FileServerOD4BProject = @()

foreach ($user in $users) {        
    $SourceFolder= $user.SourceFolder
    $importEmailAddress =  $user.UserPrincipalName 

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

            [array]$FileServerToOD4BProject += New-Object PSObject -Property @{ProjectName=$ProjectName;SourceFolder=$SourceFolder;EmailAddress=$importEmailAddress} 
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

$customerUrlId = Get-CustomerUrlId -CustomerOrganizationId $customerOrganizationId

$url = "https://manage.bittitan.com/customers/$customerUrlId/users?qp_currentWorkgroupId=$workgroupId"

Write-Host
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
    #export the project info to CSV file
    $FileServerToOD4BProject| Select-Object ProjectName,SourceFolder,EmailAddress | sort { $_.UserPrincipalName } |Export-Csv -Path $workingDir\FileServerToOD4BProject.csv -NoTypeInformation -force

    #Open the CSV file
    Start-Process -FilePath $workingDir\FileServerToOD4BProject.csv

    $msg = "SUCCESS: CSV file CSV file with the script output '$workingDir\FileServerToOD4BProject.csv' opened."
    Write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg -LogFile $logFile
    $msg = "INFO: This CSV file will be used by Start-MW_FileServerToOD4BMigrations.ps1 script to automatically submit all home directories for migration."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host
}
catch {
    $msg = "ERROR: Failed to export and open '$workingDir\FileServerToOD4BProject.csv' CSV file."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg -LogFile $logFile

##END SCRIPT