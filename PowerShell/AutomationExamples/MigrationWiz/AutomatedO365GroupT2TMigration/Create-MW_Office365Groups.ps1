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
    This script will move mailboxes from a mailbox project to a target project
	
.NOTES
	Author			For any questions contact Technical Sales Specialist Team <TSTeam@bittitan.com> or the author of this script Pablo Galan Sabugo <pablog@bittitan.com> 
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
#>

Function Import-CSV_RecipientMapping {

    Get-FileName $workingDir

    ##Import the CSV file
    Try{
        $CSVFile = Import-Csv -Path $global:inputFile
    }
    Catch [Exception] {
        Write-Host -ForegroundColor Red "ERROR: Failed to import the CSV file '$script:inputFile'."
        Write-Host -ForegroundColor Red $_.Exception.Message
        Exit
    }

    #Check if CSV is formated properly
    If (!$CSVFile.SourceEmailAddress -or !$CSVFile.DestinationEmailAddress) {
        Write-Host -ForegroundColor Red "ERROR: The CSV file format is invalid. It must have 2 columns: 'SourceEmailAddress' and 'DestinationEmailAddress' "
        Exit 
    }

    #Load existing advanced options
    $ADVOPTString += $Connector.AdvancedOptions
    $ADVOPTString += "`n"

    $count=0

    #Processing CSV into string
    Write-Host "         INFO: Applying RecipientMappings from CSV File:"

    $CSVFile | ForEach-Object {

       $sourceAddress = $_.SourceEmailAddress
       $destinationAddress = $_.DestinationEmailAddress

       $recipientMapping = "RecipientMapping=`"@$sourceAddress->@$destinationAddress`""

       $count+=1

       Write-Host -ForegroundColor Green "         SUCCESS: $recipientMapping applied." 
      
       $allRecipientMappings += $recipientMapping
       $allRecipientMappings += "`n"
    }

    Write-Host -ForegroundColor Green "         SUCCESS: CSV file '$script:inputFile' succesfully processed. $count recipient mappings applied."

    Return $allRecipientMappings
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
$logFileName = "$(Get-Date -Format yyyyMMdd)_Create-MW_Office365Groups.log"
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
$exportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "source" -EndpointType "ExchangeOnline2"
#Get source endpoint credentials
[PSObject]$exportEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $exportEndpointId 

#Create a PSCredential object to connect to source Office 365 tenant
$administrativeUsername = $exportEndpointData.AdministrativeUsername
$administrativePassword = ConvertTo-SecureString -String $($exportEndpointData.AdministrativePassword) -AsPlainText -Force
$o365Credentials = New-Object System.Management.Automation.PSCredential ($administrativeUsername, $administrativePassword)

$sourceO365Session = Connect-ExchangeOnlineSource -O365Credentials $o365Credentials 

#Select destination endpoint
$importEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "destination" -EndpointType "ExchangeOnline2"
#Get destination endpoint credentials
[PSObject]$importEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $importEndpointId 

write-host
$groups = Get-O365groups

if($groups -eq $null) {
    $msg = "INFO: No Office 365 Groups found in source Office 365 tenant. Script aborted."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Exit
}  

$MigrationWizProjectArray = @()
$alreadyCreated = $false
$totalLines = $groups.Count
$processedLines = 0

foreach ($group in $groups) {

    $groupSrcUrl = $group.srcSharePointSiteUrl
    $groupDstUrl = $group.dstSharePointSiteUrl
    $srcPrimarySMTPAddress = $group.srcPrimarySmtpAddress
    $dstPrimarySMTPAddress = $group.dstPrimarySmtpAddress

    $groupName = $group.srcDisplayName

    write-host 
    $msg = "INFO: Processing Office 365 Group '$groupName'."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile

    if($groupName -eq "" -or $groupSrcUrl -eq "" -or $groupDstUrl -eq "" -or $srcPrimarySMTPAddress  -eq "" -or $dstPrimarySMTPAddress -eq "") {
        $msg = "INFO: Skipping Office 365 Group '$groupName'. Missing data in the CSV file."
        Write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg -LogFile $logFile
        Continue  
    }
         
    #Create O365 Group source endpoint
    
    $exportEndpointName = "SRC-$srcPrimarySMTPAddress"
    $endpointTypeName = "ManagementProxy.ManagementService.SharePointConfiguration"
    $endpointType = "Office365Groups"
    $exportConfiguration = New-Object -TypeName $endpointTypeName -Property @{
        "Url" = $groupSrcUrl;
        "AdministrativeUsername" = $exportEndpointData.AdministrativeUsername;
        "AdministrativePassword" = $exportEndpointData.AdministrativePassword;
        "UseAdministrativeCredentials" = $true
    }

    $exportEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "source" -EndpointType $endpointType  -EndpointName $exportEndpointName -EndpointConfiguration $exportConfiguration

    #Create O365 Group destination endpoint

    $importEndpointName = "DST-$dstPrimarySMTPAddress"
    $endpointTypeName = "ManagementProxy.ManagementService.SharePointConfiguration"
    $endpointType = "Office365Groups"
    $importConfiguration = New-Object -TypeName $endpointTypeName -Property @{
        "Url" = $groupDstUrl;
        "AdministrativeUsername" = $importEndpointData.AdministrativeUsername;
        "AdministrativePassword" = $importEndpointData.AdministrativePassword;
        "UseAdministrativeCredentials" = $true
    }
    
    [guid]$importEndpointId = create-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "destination" -EndpointType $endpointType -EndpointName $importEndpointName -EndpointConfiguration $importConfiguration
     

#Create O365 Group Mailbox project
    
    if ($alreadyCreated -eq $false) {

        #$ProjectName = "Mailbox-$groupName"
        $ProjectName = "NGA HR-Mailbox-O365 Groups conversations"
        $projectTypeName = "MigrationProxy.WebApi.ExchangeConfiguration"
        $ProjectType = "Mailbox"
        $importType = "ExchangeOnline2"
        $exportType = "ExchangeOnline2"
        $exportEndpointId = $exportEndpointId
        $importEndpointId = $importEndpointId
        $exportConfiguration = New-Object -TypeName $projectTypeName -Property @{
            "Url" = $exportEndpointData.Url;
            "AdministrativeUsername" = $exportEndpointData.AdministrativeUsername;
            "AdministrativePassword" = $exportEndpointData.AdministrativePassword;
            "UseAdministrativeCredentials" = $true
        }
        $importConfiguration = New-Object -TypeName $projectTypeName -Property @{
            "Url" = $importEndpointData.Url;
            "AdministrativeUsername" = $importEndpointData.AdministrativeUsername;
            "AdministrativePassword" = $importEndpointData.AdministrativePassword;
            "UseAdministrativeCredentials" = $true
        }
        $folderFilter = "^(?!Inbox|Calendar)"

        if ($global:sameEmailAddresses) {

            $sourceDomain = Get-TenantDomain -Credentials $o365Credentials
            $vanityDomains = @(Get-VanityDomains -Credentials $o365Credentials)

            if(!$vanityDomains -and $global:destinationDomains.count -eq 1) {
            
                $recipientMapping += "RecipientMapping=`"@$sourceDomain->@$global:destinationDomains`""

                $msg = "INFO: Since you are migrating to the same email addresses, this '$recipientMapping' will be applied."
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile
            }
            else {
                $msg = "ACTION: Since you are migrating to the same email addresses but there are several domains, please select the RecipientMapping CSV file with SourceEmailAddress and DestinationEmailAddress."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg -LogFile $logFile  
            
                $recipientMapping = Import-CSV_RecipientMapping      
            }
        }
        elseif(!$global:sameEmailAddresses -and $script:sameUserName) {

            $sourceDomain = $srcPrimarySMTPAddress.Split("@")[1]
            $vanityDomains = @(Get-VanityDomains -Credentials $o365Credentials)

            if(!$vanityDomains -and $global:destinationDomains.count -eq 1) {

                $recipientMapping += "RecipientMapping=`"@$sourceDomain->@$global:destinationDomains`""

                $msg = "INFO: Since you are migrating to a different domain but with same email prefixes, this '$recipientMapping' will be applied."
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile
            }
            elseif($vanityDomains -and $vanityDomains.count -eq 1){
            
                $recipientMapping += "RecipientMapping=`"@$sourceDomain->@$vanityDomains`""
            
                $msg = "INFO: Since you are migrating to a different domain but with same email prefixes, this '$recipientMapping' will be applied."
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile
            }
            else {
                $msg = "ACTION: Since you are migrating to different domain but with same email prefixes and there are several domains, please select the RecipientMapping CSV file with SourceEmailAddress and DestinationEmailAddress."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg -LogFile $logFile

                $recipientMapping = Import-CSV_RecipientMapping 
            }
        }
        elseif(!$global:sameEmailAddresses -and !$global:sameUserName ) {
        
            $msg = "ACTION: Since you are migrating to different domain and email prefixes, please select the RecipientMapping CSV file with SourceEmailAddress and DestinationEmailAddress."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg -LogFile $logFile

            $recipientMapping = Import-CSV_RecipientMapping    
        }
        
        $advancedOptions = $recipientMapping

        if($totalLines -ge 400) {
            $maximumSimultaneousMigrations = 400
        }
        else {
            $maximumSimultaneousMigrations = $totalLines
        }

        $mailboxConnectorId = Create-MW_Connector -CustomerOrganizationId $customerOrganizationId `
        -ProjectName $ProjectName `
        -ProjectType $ProjectType `
        -importType $importType `
        -exportType $exportType `
        -exportEndpointId $exportEndpointId `
        -importEndpointId $importEndpointId `
        -exportConfiguration $exportConfiguration `
        -importConfiguration $importConfiguration `
        -advancedOptions $advancedOptions `
        -folderFilter $folderFilter `
        -maximumSimultaneousMigrations $maximumSimultaneousMigrations
    
        [array]$MigrationWizProjectArray += New-Object PSObject -Property @{
            ProjectName = $ProjectName;
        }

        $alreadyCreated = $true 
    }
     
     try {
        $result = Add-MW_Mailbox -ticket $global:mwTicket -ConnectorId $mailboxConnectorId -ImportEmailAddress $dstPrimarySMTPAddress -ExportEmailAddress $srcPrimarySMTPAddress 
    }
    catch {
        $msg = "ERROR: Failed to add source and destination primary SMTP address." 
        write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg -LogFile $logFile    
    }

#Create O365 Group Document project
    
    $ProjectName = "NGA HR-Document-$groupName"
    $ProjectType = "Storage"    
    $importType = "Office365Groups"
    $exportType = "Office365Groups"

    $exportEndpointId = $exportEndpointId
    $importEndpointId = $importEndpointId

    $exportTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
    $exportConfiguration = New-Object -TypeName $exportTypeName -Property @{
        "Url" = $exportEndpointData.Url;
        "AdministrativeUsername" = $exportEndpointData.AdministrativeUsername;
        "AdministrativePassword" = $exportEndpointData.AdministrativePassword;
        "UseAdministrativeCredentials" = $true
    }
    $importTypeName = "MigrationProxy.WebApi.SharePointConfiguration"
    $importConfiguration = New-Object -TypeName $importTypeName -Property @{
        "Url" = $importEndpointData.Url;
        "AdministrativeUsername" = $importEndpointData.AdministrativeUsername;
        "AdministrativePassword" = $importEndpointData.AdministrativePassword;
        "UseAdministrativeCredentials" = $true
    }

    $AdvancedOptions = "Tags=IpLockDown! InitializationTimeout=28800000 FolderLimit=20000 IncreasePathLengthLimit=1 SyncItems=1"

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
    -maximumSimultaneousMigrations 100 `
    -MaxLicensesToConsume 10
    
    try {
        $ImportLibrary = "Shared Documents"
        $ExportLibrary = "Shared Documents"
    
        $result = Add-MW_Mailbox -ticket $global:mwTicket -ConnectorId $connectorId  -ImportLibrary $ImportLibrary -ExportLibrary $ExportLibrary

        $ProcessedLines += 1
    }
    catch {
        $msg = "ERROR: Failed to add source and destination 'Shared Documents' libraries." 
        write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg -LogFile $logFile    
    }
    
    [array]$MigrationWizProjectArray += New-Object PSObject -Property @{
        ProjectName = $ProjectName;
    }      
}

write-Host
$msg = "SUCCESS: $ProcessedLines out of $totalLines Office 365 Groups have been processed." 
write-Host -ForegroundColor Green $msg
Log-Write -Message $msg -LogFile $logFile
 
try {
    $MigrationWizProjectArray| Export-Csv -Path $workingDir\O365GroupProjects.csv -NoTypeInformation -force

    $msg = "SUCCESS: CSV file '$workingDir\O365GroupProjects.csv' processed, exported and open."
    Write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg -LogFile $logFile
    $msg = "INFO: This CSV file will be used by Start-MW_Office365GroupMigrations.ps1 script to automatically submit all migrations for migration."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile
}
catch {
    $msg = "ERROR: Failed to export MigrationWiz projects to '$workingDir\O365GroupProjects.csv' CSV file."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}

try {
    #Open the CSV file
    Start-Process -FilePath $workingDir\O365GroupProjects.csv
}
catch {
    $msg = "ERROR: Failed to open '$workingDir\O365GroupProjects.csv' CSV file."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}


$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg -LogFile $logFile

##END SCRIPT