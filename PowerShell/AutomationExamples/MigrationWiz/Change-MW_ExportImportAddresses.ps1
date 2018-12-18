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
    This script changes source or destination email addresses in a MigrationWiz mailbox project.
    It will export all mailboxes from a MigrationWiz project to a CSV file with these headers: 
        ProjectName,ConnectorId,MailboxId,ExportEmailAddress,NewExportEmailAddress,ImportEmailAddress,NewImportEmailAddress
    To change source email address just edit NewExportEmailAddress value of the corresponding mailbox.
    To change the destination email address just edit NewImportEmailAddress value of the corresponding mailbox.

	
.NOTES
	Author			For any questions contact Technical Sales Specialist Team <TSTeam@bittitan.com> or the author of this script Pablo Galan Sabugo <pablog@bittitan.com> 
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. BitTitan cannot be held responsible for any misuse of the script.
	Change Log
        Please check the "Instructions" task of this Runbook for a detailed change log and full instructions
    Version: 1.1
#>

### Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory=$true)] [string]$workingDir,
        [parameter(Mandatory=$true)] [string]$logDir
    )
    if ( !(Test-Path -Path $workingDir)) {
		try {
			$suppressOutput = New-Item -ItemType Directory -Path $workingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($workingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
		}
		catch {
            $msg = "ERROR: Failed to create '$workingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
		}
    }
    if ( !(Test-Path -Path $logDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($logDir)' for log files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($logDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
}

### Function to write information to the Log File
Function Log-Write
{
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}


Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $global:inputFile = $OpenFileDialog.filename

    if($OpenFileDialog.filename -eq "") {
		    # create new import file
	        $inputFileName = "O365groups-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
            $global:inputFile = "$initialDirectory\$inputFileName"

		    $csv = "srcDisplayName,srcSharePointSiteUrl,srcPrimarySmtpAddress,dstSharePointSiteUrl,dstPrimarySmtpAddress,`r`n"
		    $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -force -value $csv

            $msg = "SUCCESS: Empty CSV file '$global:inputFile' created."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
            $msg = "WARNING: Populate the CSV file with the Office 365 groups you want to process and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$global:inputFile'."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg

		    # open file for editing
		    Start-Process $file 

		    do {
			    $confirm = (Read-Host -prompt "Are you done editing the import CSV file?  [Y]es or [N]o")
		        if($confirm -eq "Y") {
			        $importConfirm = $true
		        }

		        if($confirm -eq "N") {
			        $importConfirm = False
		        }
		    }
		    while(-not $importConfirm)
            
            $msg = "SUCCESS: CSV file '$global:inputFile' saved."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
    }
    else{
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg
    }
}

### Function to display the workgroups created by the user
Function Select-MSPC_Workgroup {

    #######################################
    # Display all mailbox workgroups
    #######################################

    $workgroupPageSize = 100
  	$workgroupOffSet = 0
	$workgroups = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC workgroups..."

    do
    {   try {
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize) #-CreatedBySystemUserId $ticket.SystemUserId 
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
    
        if($workgroupsPage) {
            $workgroups += @($workgroupsPage)
            foreach($Workgroup in $workgroupsPage) {
                Write-Progress -Activity ("Retrieving workgroups (" + $workgroups.Length + ")") -Status $Workgroup.Id
            }
            
            $workgroupOffset += $workgroupPageSize
        }

    } while($workgroupsPage)

    

    if($workgroups -ne $null -and $workgroups.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $workgroups.Length.ToString() + " Workgroup(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No workgroups found." 
        Exit
    }

    #######################################
    # Prompt for the mailbox Workgroup
    #######################################
    if($workgroups -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a Workgroup:" 
        Write-Host -Object "INFO: your default workgroup has no name, only Id." 

        for ($i=0; $i -lt $workgroups.Length; $i++)
        {
            $Workgroup = $workgroups[$i]
            if($Workgroup.Name -eq $null) {
                Write-Host -Object $i,"-",$Workgroup.Id
            }
            else {
                Write-Host -Object $i,"-",$Workgroup.Name
            }
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($workgroups.count -eq 1) {
                $msg = "INFO: There is only one workgroup. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $Workgroup=$workgroups[0]
                Return $Workgroup.Id
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($workgroups.Length-1) + ", or x")
            }
            
            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $workgroups.Length))
            {
                $Workgroup=$workgroups[$result]
                Return $Workgroup.Id
            }
        }
        while($true)

    }

}

### Function to display all customers
Function Select-MSPC_Customer {

    param 
    (      
        [parameter(Mandatory=$true)] [String]$WorkgroupId
    )

    #######################################
    # Display all mailbox customers
    #######################################

    $customerPageSize = 100
  	$customerOffSet = 0
	$customers = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC customers..."

    do
    {   
        try { 
            $customersPage = @(Get-BT_Customer -WorkgroupId $WorkgroupId -IsDeleted False -IsArchived False -PageOffset $customerOffSet -PageSize $customerPageSize)
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC customers."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
    
        if($customersPage) {
            $customers += @($customersPage)
            foreach($customer in $customersPage) {
                Write-Progress -Activity ("Retrieving customers (" + $customers.Length + ")") -Status $customer.CompanyName
            }
            
            $customerOffset += $customerPageSize
        }

    } while($customersPage)

    

    if($customers -ne $null -and $customers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $customers.Length.ToString() + " customer(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No customers found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox customer
    #######################################
    if($customers -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select a customer:" 

        for ($i=0; $i -lt $customers.Length; $i++)
        {
            $customer = $customers[$i]
            Write-Host -Object $i,"-",$customer.CompanyName
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($customers.count -eq 1) {
                $msg = "INFO: There is only one customer. Selected by default."
                Write-Host -ForegroundColor yellow  $msg
                Log-Write -Message $msg
                $customer=$customers[0]
                Return $customer.OrganizationId
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($customers.Length-1) + ", or x")
            }

            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $customers.Length))
            {
                $customer=$customers[$result]
                Return $Customer.OrganizationId
            }
        }
        while($true)

    }

}

### Function to display all mailbox connectors
Function Select-MW_Connector {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId
    )

    #######################################
    # Display all mailbox connectors
    #######################################
    
    $connectorPageSize = 100
  	$connectorOffSet = 0
	$connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving mailbox connectors ..."
    
    do
    {
        $connectorsPage = @(Get-MW_MailboxConnector -ticket $global:mwTicket -OrganizationId $customerOrganizationId -PageOffset $connectorOffSet -PageSize $connectorPageSize)
    
        if($connectorsPage) {
            $connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($connectors -ne $null -and $connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $connectors.Length.ToString() + " mailbox connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No mailbox connectors found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox connector
    #######################################
    if($connectors -ne $null)
    {
        

        for ($i=0; $i -lt $connectors.Length; $i++)
        {
            $connector = $connectors[$i]
            Write-Host -Object $i,"-",$connector.Name
        }
        Write-Host -Object "x - Exit"
        Write-Host

        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the source mailbox connector:" 

        do
        {
            $result = Read-Host -Prompt ("Select 0-" + ($connectors.Length-1) + " o x")
            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $connectors.Length))
            {
                $script:connector=$connectors[$result]
                Break
            }
        }
        while($true)
    }

}

Function Display-MW_ConnectorData {

    Write-Host
    $msg = "INFO: Retrieving migrations from '$($script:connector.Name)' project..."
    Write-Host $msg
    Log-Write -Message $msg

    $mailboxes = @()
    $mailboxesArray = @()

        # Retrieve all mailboxes from the specified project
        $mailboxOffSet = 0
        $mailboxPageSize = 100
        $mailboxes = $null

        do {
            $mailboxesPage = @(Get-MW_Mailbox -Ticket $global:mwTicket -FilterBy_Guid_ConnectorId $script:connector.Id -PageOffset $mailboxOffSet -PageSize $mailboxPageSize) | sort { $_.ExportEmailAddress.length }

            if($mailboxesPage) {
                $mailboxes += @($mailboxesPage)

                foreach($mailbox in $mailboxesPage) {
                    if($script:connector.ProjectType -eq "Mailbox") {
                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz mailbox project") -Status $mailbox.ExportEmailAddress.ToLower()

                        $tab = [char]9
                        Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                        Write-Host -nonewline "$($script:connector.Name) "               
                        write-host -nonewline -ForegroundColor Yellow "ExportEmailAddress: "
                        write-host -nonewline "$($mailbox.ExportEmailAddress)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportEMailAddress: "
                        write-host "$($mailbox.ImportEmailAddress)"

                        $mailboxLineItem = New-Object PSObject
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ConnectorId -Value $script:connector.Id
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name MailboxId -Value $mailbox.Id
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportEmailAddress -Value $mailbox.ExportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportEmailAddress -Value $mailbox.ExportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportEmailAddress -Value $mailbox.ImportEmailAddress
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewImportEmailAddress -Value $mailbox.ImportEmailAddress

                        $mailboxesArray += $mailboxLineItem
                    }
                    elseif($script:connector.ProjectType -eq "Storage") {
                        Write-Progress -Activity ("Retrieving migrations for '$($script:connector.Name)' MigrationWiz Office 365 Groups project") -Status $mailbox.ExportLibrary.ToLower()

                        $tab = [char]9
                        Write-Host -nonewline -ForegroundColor Yellow  "Project: "
                        Write-Host -nonewline "$($script:connector.Name) "               
                        write-host -nonewline -ForegroundColor Yellow "ExportLibrary: "
                        write-host -nonewline "$($mailbox.ExportLibrary)$tab"
                        write-host -nonewline -ForegroundColor Yellow "ImportLibrary: "
                        write-host "$($mailbox.ImportLibrary)"

                        $mailboxLineItem = New-Object PSObject
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ProjectName -Value $script:connector.Name
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ExportLibrary -Value $mailbox.ExportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewExportLibrary -Value $mailbox.ExportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name ImportLibrary -Value $mailbox.ImportLibrary
                        $mailboxLineItem | Add-Member -MemberType NoteProperty -Name NewImportLibrary -Value $mailbox.ImportLibrary

                        $mailboxesArray += $mailboxLineItem
                    }

                }

                $mailboxOffSet += $mailboxPageSize
            }
        } while($mailboxesPage)

        if($mailboxes -ne $null -and $mailboxes.Length -ge 1) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $($mailboxes.Length) migrations found." 
        }
        else {
            Write-Host -ForegroundColor Red "INFO: No migrations found. Script aborted." 
            Exit
        }

        try {
            $mailboxesArray | Export-Csv -Path $workingDir\MailboxData.csv -NoTypeInformation -force

            $msg = "SUCCESS: CSV file '$workingDir\MailboxData.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
        catch {
            $msg = "ERROR: Failed to export mailboxes to '$workingDir\MailboxData.csv' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }

        try {
            #Open the CSV file for editing
            Start-Process -FilePath $workingDir\MailboxData.csv
        }
        catch {
            $msg = "ERROR: Failed to open '$workingDir\MailboxData.csv' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Exit
        }
    


}

Function Change-MW_ExportImportAddresses {


	if (Test-Path $workingDir) {

        $migrations = @(Import-Csv -Path $workingDir\MailboxData.csv) 
		$migrations | ForEach-Object {

            $mailbox = Get-MW_Mailbox -Ticket $global:mwTicket -ConnectorId $_.ConnectorId -Id $_.MailboxId -ImportEmailAddress $_.ImportEmailAddress -ExportEmailAddres $_.ExportEmailAddress -ErrorAction SilentlyContinue

            if ($mailbox) {
    			
                    $result = Set-MW_Mailbox -Ticket $global:mwTicket -ConnectorId $_.ConnectorId  -mailbox $mailbox -ImportEmailAddress $_.NewImportEmailAddress -ExportEmailAddres $_.NewExportEmailAddress

                    if($_.ExportEmailAddress -ne $_.NewExportEmailAddress -or $_.ImportEmailAddress -ne $_.NewImportEmailAddress) {

	                    Write-Host -NoNewline -ForegroundColor Green "[SUCCESS] "
    
                        if($_.ExportEmailAddress -ne $_.NewExportEmailAddress) {
                            $msg = "ExportEmailAddress: $($_.ExportEmailAddress) renamed to $($_.NewExportEmailAddress)$tab"
                            Write-Host -NoNewline -ForegroundColor Green $msg 
                            Log-Write -Message $msg
                        }
                        if($_.ImportEmailAddress -ne $_.NewImportEmailAddress) {
                            $msg = "ImportEmailAddress: $($_.ImportEmailAddress) renamed to $($_.NewImportEmailAddress)."
	                        Write-Host -NoNewline -ForegroundColor Green $msg 
                            Log-Write -Message $msg
                        }
                        Write-Host "`r"
                    }
                    else {
                        $msg = "[INFO] Migration ExportEmailAddress: $($_.ImportEmailAddress) $tab ImportEmailAddress: $($_.NewImportEmailAddress) has not been renamed."
                        Write-Host $msg 
                        Log-Write -Message $msg
                    }
            }else {
                $msg = "[ERROR] Failed to set Migration ExportEmailAddress: $($_.ImportEmailAddress) $tab ImportEmailAddress: $($_.NewImportEmailAddress)."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
            }

        }
	}
	else {
		Write-Host -ForegroundColor Red "ERROR: The CSV file '$workingDir\MailboxData.csv' was not found." 
	}
}

### Function to wait for the user to press any key to continue
Function WaitForKeyPress{
    $msg = "ACTION: If you have edited and saved the CSV file then press any key to continue." 
    Write-Host $msg
    Log-Write -Message $msg
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################

#Working Directory
$workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Move-MW_Mailboxes.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg

# Authenticate
$creds = Get-Credential -Message "Enter BitTitan credentials"
try {
    # Get a ticket and set it as default
    $ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan -SetDefault
    # Get a MW ticket
    $global:mwTicket = Get-MW_Ticket -Credentials $creds 
} catch {
    $msg = "ERROR: Failed to create ticket."
    Write-Host -ForegroundColor Red  $msg
    Log-Write -Message $msg
    Write-Host -ForegroundColor Red $_.Exception.Message
    Log-Write -Message $_.Exception.Message    

    Exit
}

#Select workgroup
$WorkgroupId = Select-MSPC_WorkGroup

#Select customer
$customerOrganizationId = Select-MSPC_Customer -Workgroup $WorkgroupId

#Select connector
Select-MW_Connector -customerOrganizationId $customerOrganizationId 
$result = Display-MW_ConnectorData
Write-Host

WaitForKeyPress

Write-Host
Change-MW_ExportImportAddresses

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

##END SCRIPT