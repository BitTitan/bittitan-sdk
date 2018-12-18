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
    This script will download and launch the Uploader agent from the BitTitan server, will create an Azure blob container 
    in case it does not exist in the Azure storage account and run the agent with the correct parameters to upload all the 
    File Server Home Directories to the Azure blob container.
	
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
$logFileName = "$(Get-Date -Format yyyyMMdd)_Upload-UW_FileServerToAzureBlobContainer.log"
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

$url = "https://api.bittitan.com/secure/downloads/UploaderWiz.zip"   
$outFile = "$PSScriptRoot\UploaderWiz.zip" 
$path = "$PSScriptRoot\UploaderWiz"

$checkPath = Test-Path $outFile 
if($checkPath) {
    $lastWriteTime = (get-Item -Path $path).LastWriteTime

    do {
        $confirm = (Read-Host -prompt "UploaderWiz was downloaded on $lastWriteTime. Do you want to download it again?  [Y]es or [N]o")

        if($confirm.ToLower() -eq "y") {
            $downloadUploaderWiz = $true
        }

    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
}else {
    $downloadUploaderWiz = $true
}

if($downloadUploaderWiz) {
    
    Download-File -Url $url -OutFile $outFile
}
Write-Host

$msg = "INFO: Connecting to MSPComplete to retrieve the 2 endpoints (AzureSubscription, AzureFileSystem) to connect to Azure."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

Connect-BitTitan

#Select workgroup
$WorkgroupId = Select-MSPC_WorkGroup

#Select customer
$customerOrganizationId = Select-MSPC_Customer -Workgroup $WorkgroupId

#Select source endpoint
$azureSubscriptionEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -EndpointType "AzureSubscription"
#Get source endpoint credentials
[PSObject]$azureSubscriptionEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $azureSubscriptionEndpointId 

#Create a PSCredential object to connect to Azure Active Directory tenant
$administrativeUsername = $azureSubscriptionEndpointData.AdministrativeUsername
$administrativePassword = ConvertTo-SecureString -String $($azureSubscriptionEndpointData.AdministrativePassword) -AsPlainText -Force
$azureCredentials = New-Object System.Management.Automation.PSCredential ($administrativeUsername, $administrativePassword)

#Select source endpoint
$exportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "source" -EndpointType "AzureFileSystem"
#Get source endpoint credentials
[PSObject]$exportEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $exportEndpointId 

Write-Host
# AzureRM module installation
Check-AzureRM
# Azure log in
Connect-Azure -AzureCredentials $azureCredentials -SubscriptionID $azureSubscriptionEndpointData.SubscriptionID
#Azure storage account
$storageAccount = Check-StorageAccount -StorageAccountName $exportEndpointData.AdministrativeUsername
# Azure blob container
if(!$exportEndpointData.ContainerName){
    $container = "migrationwiz"
}
$result = Check-BlobContainer -BlobContainerName "migrationwiz" -StorageAccount $storageAccount

if(!$result) {
    Create-BlobContainer -BlobContainerName $container -StorageAccount $storageAccount
}

Write-Host
do {
    Write-host -ForegroundColor Yellow  "ACTION: Enter the folder path to the FileServer root folder: "  -NoNewline
    $rootPath = Read-Host
    $rootPath = "`'$rootPath`'"

} while($rootPath -eq "")

$uploaderwizCommand = ".\UploaderWiz.exe -type azureblobs -accesskey " + $CH34 + $exportEndpointData.AdministrativeUsername + $CH34 + " -secretkey " + $CH34 + $exportEndpointData.AccessKey + $CH34 +" -container "+ $container + " -rootPath " + $CH34 + $rootpath + $CH34 + " -force True"

#Run the UploaderWiz command with with parameters
Write-Host
$msg = "INFO: Launching UploaderWiz with these parameters:`r`n$uploaderwizCommand"
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile

Write-Host
$msg = "INFO: Changing to directory '$path'."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile
cd $path

Invoke-Expression $uploaderwizCommand 

$msg = "INFO: Going back to parent directory."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile
cd ..

$msg = "INFO: UploaderWiz log file in folder '$Env:temp\UploaderWiz'."
Write-Host $msg
Log-Write -Message $msg -LogFile $logFile
#Open the CSV file
try {    
    Start-Process -FilePath "$Env:temp\UploaderWiz"

    $msg = "SUCCESS: Folder '$Env:temp\UploaderWiz' opened."
    Write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg -LogFile $logFile
}
catch {
    $msg = "ERROR: Failed to open folder '$Env:temp\UploaderWiz."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg -LogFile $logFile

##END SCRIPT
