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
    This script will create an Azure blob container in case it does not exist and a .bat file that will be sent automatically to all end users.
    as an email attachment. The email i sent from Office 365. The .bat file will have to be clicked by each end user to automatically donwload 
    the UploaderWiz agent from BitTitan server, disconnect all PST files from the Outlook profile and discover and upload all PST files to the 
    Azure blob container. The end users' email adddreses are read from a CSV file.
	
.NOTES
	Author			For any questions contact Technical Sales Specialist Team <TSTeam@bittitan.com> or the author of this script Pablo Galan Sabugo <pablog@bittitan.com> 
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    BitTitan cannot be held responsible for any misuse of the script.
	Change Log
        Please check the "Instructions" task of this Runbook for a detailed change log and full instructions
    Version: 1.1
#>

Function Get-CsvFile {
    Write-Host
    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import the user email addresses."
    Get-FileName $workingDir

    # Import CSV and validate if headers are according the requirements
    try {
        $lines = @(Import-Csv $global:inputFile)
    }
    catch {
        $msg = "ERROR: Failed to import '$inputFile' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit   
    }

    # Validate if CSV file is empty
    if ( $lines.count -eq 0 ) {
        $msg = "ERROR: '$inputFile' CSV file exist but it is empty. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Exit
    }

    # Validate CSV Headers
    $CSVHeaders = "UserEmailAddress,FirstName"
    foreach ($header in $CSVHeaders) {
        if ($lines.$header -eq "" ) {
            $msg = "ERROR: '$inputFile' CSV file does not have all the required columns. Required columns are: '$($CSVHeaders -join "', '")'. Script aborted."
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
$logFileName = "$(Get-Date -Format yyyyMMdd)_Upload-UW_PstToAzureBlobContainer.log"
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

$msg = "INFO: Connecting to MSPComplete to retrieve the 2 endpoints (AzureSubscription, Pst) to connect to Azure."
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
$exportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "source" -EndpointType "Pst"
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
    $container = "migrationwizpst"
}
$result = Check-BlobContainer -BlobContainerName "migrationwizpst" -StorageAccount $storageAccount

if(!$result) {
    Create-BlobContainer -BlobContainerName $container -StorageAccount $storageAccount
}

$disconnectPSTs = "powershell `"Write-host 'Disconnecting all PST files from your Outlook.';`$Outlook = New-Object -ComObject Outlook.Application;`$Namespace = `$Outlook.getNamespace('MAPI');`$all_psts = `$Namespace.Stores | Where-Object {(`$_.ExchangeStoreType -eq '3') -and (`$_.FilePath -like '*.pst') -and (`$_.IsDataFileStore -eq `$true)}; ForEach (`$pst in `$all_psts){`$Outlook.Session.RemoveStore(`$pst.GetRootFolder());}`" "

$downloadUploaderWiz = 'bitsadmin /transfer PST_Migration /download /priority normal "https://api.bittitan.com/secure/downloads/UploaderWiz.zip" "c:\BitTitan\UploaderWiz.zip"'

$unzipUploaderWiz = "powershell Expand-Archive c:\BitTitan\UploaderWiz.zip  -DestinationPath c:\BitTitan\UploaderWiz\ -force" 

$uploaderwizCommand = "UploaderWiz.exe -type azureblobs -accesskey " + $CH34 + $exportEndpointData.AdministrativeUsername + $CH34 + " -secretkey " + $CH34 + $exportEndpointData.AccessKey + $CH34 +" -container "+ $container + " -autodiscover true -interactive false -filefilter " + $CH34 + " *.pst" + $CH34 + " -force True"

$startUploaderWiz = "START C:\BitTitan\UploaderWiz\$uploaderwizCommand"

$batchFileCode = "@echo off`r`n`r`n$disconnectPSTs`r`n`r`n$downloadUploaderWiz`r`n`r`n$unzipUploaderWiz`r`n`r`n$startUploaderWiz`r`n"

$batFile = "C:\scripts\Migrate_PST_Files.bat"
Set-Content -Path $batFile -Value $batchFileCode -Encoding ASCII

# Azure FileShare
$result = Check-BlobContainer -BlobContainerName "batchfile" -StorageAccount $storageAccount

if(!$result) {
    Create-BlobContainer -BlobContainerName "batchfile" -StorageAccount $storageAccount -PermissionsOff  $true 
}

# upload the batch file
$result = Set-AzureStorageBlobContent -File $batFile -Container "batchfile" -Blob "migratePstFiles.bat" -Context $storageAccount.context -force
 
$url = Create-SASToken -BlobContainerName "batchfile" -BlobName "migratePstFiles.bat" -StorageAccount $storageAccount

Write-Host
$msg = "SUCCESS: Batch file '$batFile' created to be sent to all end users for them to manually run it for PST file automated migration."
Write-Host  -ForegroundColor Green $msg
Log-Write -Message $msg -LogFile $logFile
write-host
$msg = "++++++++++++++++++++++++++++++++++++++++ BATCH FILE: Migrate_PST_files.bat  ++++++++++++++++++++++++++++++++++++++++`n"
write-host $msg
Log-Write -Message $msg -LogFile $logFile

write-host $batchFileCode
Log-Write -Message $batchFileCode -LogFile $logFile

$msg = "++++++++++++++++++++++++++++++++++++++++++++++++ END BATCH FILE ++++++++++++++++++++++++++++++++++++++++++++++++++++`n"
write-host $msg
Log-Write -Message $msg -LogFile $logFile

$applyCustomFolderMapping = $false
do {
    $confirm = (Read-Host -prompt "Do you want to send the .bat file to all your users?  [Y]es or [N]o").trim()

    if($confirm.ToLower() -eq "y") {
        $sendBatchFile = $true        
    }

} while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

if($sendBatchFile) {

    $users = Get-CsvFile
    Write-Host
    $msg = "ACTION: Provide your Office 365 admin credentials to send the emails."
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg -LogFile $logFile

    $smtpCreds = Get-Credential -Message "Enter Office 365 credentials"

    Write-Host
    foreach ($user in $users) {
        
        $msg = "INFO: Sending email with .bat file to '$($user.userEmailAddress)'."
        Write-Host $msg
        Log-Write -Message $msg -LogFile $logFile

        #################################################################################
        $smtpServer = "smtp.office365.com"        
        $emailTo = $user.userEmailAddress
        $emailFrom = $smtpCreds.UserName
        $Subject = "Action required: Install the BitTitan UploaderWiz Agent on your computer."

        $body += "<tbody>"
        $body += "<center>"
        $body += "<table>"
        $body += "<tr>"
        $body += "<td align='left' valign='top'>"
        $body += "<p class='x_Logo'><a href='http://www.bittitan.com' target='_blank' rel='noopener noreferrer' data-auth='NotApplicable' title='BitTitan'><img data-imagetype='External' src='https://static.bittitan.com/Images/MSPC/MSPC_banner.png' width='600' height='50' class='x_LogoImg' alt='BitTitan' border='0'> </a></p>"
        $body += "<span style='font-family: Arial, Helvetica, sans-serif, serif, EmojiFont; font-size: 12px; color: rgb(10, 10, 10);'>"
        $body += "<p>Hello $($user.FirstName),</p>"
        $body += "<h3>Important Announcement</h3>"
        $body += "<p>We are currently planning a series of updates and improvements to our IT Services.</p>"
        $body += "<p>We are committed to creating and maintaining the best user experience with these changes. In order to do so, </br>" 
        $body += "so we will need to install an application (the BitTitan UploaderWiz Agent) that will disconnect all your PST files</br>"
        $body += "from you Outlook client and migrate them to your new Office 365 mailbox.</p>"
        $body += "<h3>Actions Required</h3>"
        $body += "<p>Complete the application installation by following these steps:</p>"
        $body += "<ol>"
        $body += "<li>Click on this link: <a href='$url' target='_blank' rel='noopener noreferrer' data-auth='NotApplicable' title='Install BitTitan UploaderWiz Agent'>Install BitTitan UploaderWiz Agent Application</a> `</li><li>Select Run. `</li></ol>"
        $body += "<p>The application will silently install. It will not impact any other work that you are doing.</p>"
        $body += "<p>We will contact you in the coming weeks if we determine that your computer requires any necessary updates.</p>"
        $body += "<hr>"
        $body += "<p>Thank you, $CustomerName IT Department</p>"
        $body += "</span></td>"
        $body += "</tr>"
        $body += "</table>"
        $body += "</td>"
        $body += "</tr>"
        $body += "</center>"
        $body += "</tbody>"
        
        #$attachment = "c:\scripts\Migrate_PST_files.bat.rename"

        #################################################################################

        try {

            $result = Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer -Port 587 -Credential $smtpCreds -UseSsl #-Attachments $attachment 

            if ($error[0].ToString() -match "Spam abuse detected from IP range.") { 
                #5.7.501 Access denied, spam abuse detected. The sending account has been banned due to detected spam activity. 
                #For details, see Fix email delivery issues for error code 451 5.7.500-699 (ASxxx) in Office 365.
                #https://support.office.com/en-us/article/fix-email-delivery-issues-for-error-code-451-4-7-500-699-asxxx-in-office-365-51356082-9fef-4639-a18a-fc7c5beae0c8 
                $msg = "      ERROR: Failed to send email to user '$emailTo'. Access denied, spam abuse detected. The sending account has been banned. "
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg -LogFile $logFile
            }
            else {
                $msg = "SUCCESS: Email with .bat file sent to end user '$emailTo'"
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg -LogFile $logFile 
           }

        }
        catch {
            $msg = "ERROR: Failed to send email to user '$emailTo'."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg -LogFile $logFile
            Log-Write -Message $_.Exception.Message -LogFile $logFile
        }
    }
}

try {
    #Open the folder containing the .bat file
    Start-Process -FilePath C:\scripts\
}
catch {
    $msg = "ERROR: Failed to open 'C:\scripts\Migrate_PST_files.bat' batch file."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg -LogFile $logFile
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message -LogFile $logFile
    Exit
}

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg -LogFile $logFile

##END SCRIPT
