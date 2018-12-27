<#
This PowerShell Script will create a new MigrationWiz Project based on 2 diffrent CSV Files
These Files has to be filled out before using this script.
The Files must uploaded before to Azure, using UploaderWiz

Remarks: This Powershell Script will only run, if you have enabled API in your profile of MIgrationWiz

.Notes

    .Version          2.0
    Author            Hans Brender/Jethro Seghers
    Creation Date     12/12/15
	Purpose/Changes   Implementing MSPC Endpoint
					  implemeting TRY / CATCH for Error Handling
                  

Required Files in this subdirectory

    ProjectImport.ps1           Reads the two csv files and produce a new Project in MigrationWiz
    Log_Functions.ps1        Functions to write a log File
    AuthorizationSettings.csv   This is the CSV, which hold all the information of the Project


#>
$Version = 3.0

Import-Module “C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll”

#Remarks: This Powershell Script will only run, if you have enabled API in your profile of MIgrationWiz
Clear Host
#############################################################################################################################################################################
# parameters, to be changed here in script:
$AuthSet="\AuthorizationSettings.csv"                                   # This is the CSV, which hold all the information of the Project


#all other varabled after this line should not be changed
##############################################################################################################################################################################
$Log="\FS2ODFB.log"                                                     # This is the logFile for the Powershell Scrpts

# including a separate Function-File for logging
$directorypath = Split-Path $script:MyInvocation.MyCommand.Path
#$directorypath = "C:\scripts"

$incFunctions = $directorypath + "\Log_Functions.ps1"

$logfile = $directorypath + $log
. $incFunctions                                                         # use the . syntax to include the functions file 


#Split the diffrent filenames
$AuthorizationSettingsCSV=$directorypath + $AuthSet


write-host $AuthorizationSettingsCSV

Log-Write -LogPath $logfile -LineValue "Version:$Version"
# Starting Log, Write 2 Lines into Log
Log-Write -LogPath $logfile -LineValue " "
$msg = "Start of Script ProjectImport.ps1"
Log-Write -LogPath $logfile -LineValue $msg
Write-Host $msg

write-host "Version: $Version"
Write-Host "Please insert your MigrationWiz Credentials"
$msg = "User inputs MigrationWiz Credentials"
Log-Write -LogPath $logfile -LineValue $msg

$cred = Get-Credential -Message "Please insert your MigrationWiz Credentials" 

Try
    {
    $ticket = Get-MW_Ticket -Credentials $cred
    }
Catch
    {
    Write-host "Wrong Credentials" -ForegroundColor Red
    $msg = $Error[0].Exception.Message
    Write-host $msg -ForegroundColor Red
    Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True   
    }


Try
    {
    $BTT = Get-BT_Ticket -Credentials $cred -ServiceType BitTitan 
    }
Catch
    {
    Write-host "Wrong Credentials" -ForegroundColor Red
    $msg = $Error[0].Exception.Message
    Write-host $msg -ForegroundColor Red
    Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True   
    }    


$msg = "Reading both csv's... and check Entries"
Write-Host $msg -ForeGroundcolor Yellow
Log-Write -LogPath $logfile -LineValue $msg


$azureConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration'
$oneDriveProConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration'
$MWazureConfiguration = New-Object -TypeName           'MigrationProxy.WebApi.AzureConfiguration'
$MWoneDriveProConfiguration = New-Object -TypeName     'MigrationProxy.WebApi.SharePointOnlineConfiguration'

#Reading the Configuration File with the different Parameter
#
# $projectName                                             Project Name in MigrationWiz
# $azureConfiguration.AdministrativeUsername               Azure Blob Storage Name
# $azureConfiguration.AccessKey                            SecretKey to access Blob-Storage in Azure
# $azureConfiguration.ContainerName                        Specifies the Container Name in Azure Blob
# $oneDriveProConfiguration.AdministrativeUsername         Global Admin in Office 365
# $oneDriveProConfiguration.AdministrativePassword         Password of the Global Admin in Office 365 
# FolderMapping                                            Entry in AuthorizationSettings.csv : if entry is YES, we will insert FolderMappings in Advanced Options     


TRY
    {
        $CLines = Import-Csv $AuthorizationSettingsCSV
        
        foreach ($cLine in $Clines)
        {
             #Define Project name based on Email address
            
            $ProjectName = $CLine.DestinationEmail + "_PS"
            write-host $ProjectName
            If ($ProjectName.Length -gt 0)
            {
                Log-Write -LogPath $logfile -LineValue "Found ProjectName: $ProjectName"
                Write-Host "Found ProjectName: " $ProjectName 
            }
            else
            {
                $msg ="Found no ProjectName in AuthorizationSettings.csv "
                Write-host $msg -ForegroundColor Red
                Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True
            }
        

            $SecretKey=$cLine.SecretKey                                          
            If ($secretKey.Length -gt 0)
            {
                $azureConfiguration.AccessKey =$SecretKey
                Log-Write -LogPath $logfile -LineValue "Found SecretKey: $SecretKey"
                Write-Host "Found SecretKey: " $SecretKey  
            }
            else
            {
                $msg ="Found no SeceretKey in AuthorizationSettings.csv "
                Write-host $msg -ForegroundColor Red
                Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True
            }

            $SourceFolder = $cLine.SourceFolder

            $ContainerName= Rename-AzureContainer $SourceFolder

            If ($ContainerName.Length -gt 0)
            {
                $azureConfiguration.ContainerName =$ContainerName
                Log-Write -LogPath $logfile -LineValue "Found SecretKey: $ContainerName"
                Write-Host "Found ContainerName: " $ContainerName            
            }
            else
            {
                $msg ="Found no ConatinerName in AuthorizationSettings.csv "
                Write-host $msg -ForegroundColor Red
                Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True
            }

                
            $AccessKey =$cLine.AccessKey
            Write-Host $AccessKey
            If ($AccessKey.Length -gt 0)
            {
                $azureConfiguration.AdministrativeUsername =$AccessKey
                Log-Write -LogPath $logfile -LineValue "Found AccessKey: $Accesskey"
                Write-Host "Found AccessKey: " $AccessKey
            }
            else
            {
                $msg ="Found no AccessKey in AuthorizationSettings.csv "
                Write-host $msg -ForegroundColor Red
                Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True
            }
                

            $ODFBAdmin =$Cline.ODFBAdmin
            If ($ODFBAdmin.Length -gt 0)
            {
                $oneDriveProConfiguration.AdministrativeUsername =$ODFBAdmin
                Log-Write -LogPath $logfile -LineValue "Found ODFB Admin: $ODFBAdmin"
                Write-Host "Found ODFB Admin: " $ODFBAdmin
            }
            else
            {
                $msg ="Found no ODFB Admin in AuthorizationSettings.csv "
                Write-host $msg -ForegroundColor Red
                Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True
            }

            $ODFBPwd= $Cline.ODFBPwd
            If ($ODFBPwd.Length -gt 0)
            {
                $oneDriveProConfiguration.AdministrativePassword  =$ODFBPwd
                Log-Write -LogPath $logfile -LineValue "Found ODFB Admin Pwd: $ODFBPwd"
                Write-Host "Found ODFB Admin Pwd: " $ODFBPwd
            }
            else
            {
                $msg ="Found no ODFB Admin Pwd in AuthorizationSettings.csv "
                Write-host $msg -ForegroundColor Red
                Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True
            }
               
            $BitTitanDC="Canada","NorthEurope","NorthAmerica","WesternEurope","AsiaPacific","Australia","Japan","SouthAmerica"                              
            $LZone=$Cline.BitTitanDatacenter 
            #FIX
                                                                         
            If ($BitTitanDC -eq $LZone)
            {
                $Zone = $LZone
                Log-Write -LogPath $logfile -LineValue "Found BitTitan DataCenter: $LZone"
                Write-Host "Found BitTitan DataCenter: " $LZone
            }
            Else
            {
                $msg ="Found no valid BitTitan DataCenter in AuthorizationSettings.csv  Found: " +$LZone
                Write-host $msg -ForegroundColor Red
                Log-Write -LogPath $logfile -LineValue $msg
                $zone = $BitTitanDC[0]  #replacing wrong or invalid entry wityh 'NorthAmerica'
                Log-Write -LogPath $logfile -LineValue "Replace invalid BitTitan DataCenter: $LZone with $zone "
                Write-Host "Replace invalid BitTitan DataCenter: $LZone with $zone "

            }
            
            $azureConfiguration.UseAdministrativeCredentials = 1  
            $oneDriveProConfiguration.UseAdministrativeCredentials = 1   
            
            #Read Customer From CSV
            $customerName = $cline.CustomerName
            
            #Make project 

            TRY
            {
                $customer = Get-BT_Customer -Ticket $BTT -Environment BT -FilterBy_String_CompanyName $customerName
                write-host "Found Customer"

                $t = Get-BT_Ticket -Credentials $cred -ServiceType BitTitan -OrganizationId $customer.OrganizationId
                
                $MWT = Get-MW_Ticket -Credentials $cred 
                
                $oneDriveProConfiguration.AzureStorageAccountName = $azureConfiguration.AdministrativeUsername
                $oneDriveProConfiguration.AzureAccountKey = $azureConfiguration.AccessKey

                $destinationEndpointName = "OD4B v2 PS"
                $destinationEndpoint = Get-BT_Endpoint -Ticket $t -FilterBy_String_Name $destinationEndpointName -FilterBy_Boolean_IsDeleted false              
                
                if (!$destinationEndpoint)
                {
                    Log-Write -LogPath $logfile -LineValue "Creating Endpoint $destinationEndpointName"
                    $destinationEndpoint = Add-BT_Endpoint -Ticket $t -Name $destinationEndpointName -Type "OneDriveProAPI" -Configuration $oneDriveProConfiguration
                    Log-Write -LogPath $logfile -LineValue "Endpoint $destinationEndpointName created"
                }
                else
                {
                    Log-Write -LogPath $logfile -LineValue "Found Endpoint $destinationEndpointName"
                }

                $sourceEndpointName = "HomeDrive PS"
                $sourceEndpoint = Get-BT_Endpoint -Ticket $t -FilterBy_String_Name $sourceEndpointName -FilterBy_Boolean_IsDeleted false
                if (!$sourceEndpoint)
                {
                    Log-Write -LogPath $logfile -LineValue "Creating Endpoint $sourceEndpointName"
                    $sourceEndpoint = Add-BT_Endpoint -Ticket $t -Name $sourceEndpointName -Type "AzureFileSystem" -Configuration $azureConfiguration
                    Log-Write -LogPath $logfile -LineValue "Endpoint $sourceEndpointName created"                 
                }
                else
                {
                    Log-Write -LogPath $logfile -LineValue "Found Endpoint $sourceEndpointName" 
                }

                $MWazureConfiguration = New-Object -TypeName           'MigrationProxy.WebApi.AzureConfiguration'
                $MWoneDriveProConfiguration = New-Object -TypeName     'MigrationProxy.WebApi.SharePointOnlineConfiguration'

                $MWazureConfiguration.AccessKey = $azureConfiguration.AccessKey
                $MWazureConfiguration.AdministrativeUsername = $azureConfiguration.AdministrativeUsername
                $MWazureConfiguration.UseAdministrativeCredentials = $azureConfiguration.UseAdministrativeCredentials
                $MWazureConfiguration.ContainerName = $azureConfiguration.ContainerName

                $MWoneDriveProConfiguration.AdministrativePassword = $oneDriveProConfiguration.AdministrativePassword
                $MWoneDriveProConfiguration.AdministrativeUsername = $oneDriveProConfiguration.AdministrativeUsername
                $MWoneDriveProConfiguration.AzureAccountKey = $oneDriveProConfiguration.AzureAccountKey
                $MWoneDriveProConfiguration.AzureStorageAccountName = $oneDriveProConfiguration.AzureStorageAccountName
                $MWoneDriveProConfiguration.UseAdministrativeCredentials = $oneDriveProConfiguration.UseAdministrativeCredentials

                
                $connector = Add-MW_MailboxConnector -ticket $MWT -Name $ProjectName -ProjectType Storage `
                        -ImportType OneDriveProAPI -ImportConfiguration $MWoneDriveProConfiguration `
                        -ExportType AzureFileSystem -ExportConfiguration $MWazureConfiguration `
                        -OrganizationId $customer.OrganizationId `
                        -SelectedExportEndpointId $sourceEndpoint.Id `
                        -SelectedImportEndpointId $destinationEndpoint.Id `
                        -UserId $MWT.UserId -MaximumDataTransferRate ([int]::MaxValue) -MaximumDataTransferRateDuration 600000  `
                        -MaximumSimultaneousMigrations 100 -PurgePeriod 90 -MaximumItemFailures 100 -ZoneRequirement $zone -ItemEndDate ((Get-Date).AddYears(5)) -MaxLicensesToConsume 5 -AdvancedOptions "RenameConflictingFiles=1 ShrinkFoldersMaxLength=200"

                write-host "Project is created"

                #Add User
                $result = Add-MW_Mailbox -ticket $MWT -ImportEmailAddress $CLine.DestinationEmail -ConnectorId $Connector.Id

                $verification = Add-MW_MailboxMigration -Ticket $MWT -MailboxId $result.Id -Type Verification -Status Submitted -ConnectorId $connector.Id -UserId $MWT.UserId -Priority 1
                
            }
            catch
            {
                Write-host "There was an Error building the connector" -ForegroundColor Red
                $msg = $Error[0].Exception.Message
                Write-host $msg -ForegroundColor Red
                Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True

            }

            

        }        
    }
catch
    {
    $msg = $Error[0].Exception.Message
    Write-host $msg -ForegroundColor Red
    Log-Error -LogPath $logfile -ErrorDesc  $msg -ExitGracefully $True
    }

   