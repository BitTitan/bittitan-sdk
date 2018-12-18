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
    This file contains only functions 
	
.NOTES
	Author			For any questions contact Technical Sales Specialist Team <TSTeam@bittitan.com> or the author of this script Pablo Galan Sabugo <pablog@bittitan.com> 
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
#>


#Function to authenticate to BitTitan SDK
Function Connect-BitTitan {
    [CmdletBinding()]
    # Authenticate
    $creds = Get-Credential -Message "Enter BitTitan credentials"
    # Get a ticket and set it as default
    $ticket = Get-BT_Ticket -Credentials $creds -SetDefault -ServiceType BitTitan -ErrorAction SilentlyContinue
    # Get a MW ticket
    $global:mwTicket = Get-MW_Ticket -Credentials $creds -ErrorAction SilentlyContinue

    if(!$ticket -or !$global:mwTicket) {
        $msg = "ERROR: Failed to create ticket. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host
        Write-Host -ForegroundColor Red $error[0]
        Log-Write -Message $error[0]  

        Exit
    }
}

# Function to create the working and log directories
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

# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message,
        [Parameter(Mandatory=$true)]    [string]$LogFile
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}

# Function to create source EXO PowerShell session
Function Connect-ExchangeOnlineSource {

    param 
    (      
        [parameter(Mandatory=$false)] [System.Management.Automation.PSCredential]$O365Credentials
    )
    
    #Prompt for source Office 365 global admin Credentials
    $msg = "INFO: Connecting to the source Office 365 tenant."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile

    if (!($SourceO365Session.State)) {
        try {
            $loginAttempts = 0
            do {
                $loginAttempts++

                # Connect to Source Exchange Online
                if(!$O365Credentials) {
                    $SourceO365Creds = Get-Credential -Message "Enter Your Source Office 365 Admin Credentials."
                    if (!($SourceO365Creds)) {
                        $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg -LogFile $logFile
                        Exit
                    }
                }
                else {
                    $SourceO365Creds = Get-Credential $O365Credentials
                }
                $SourceO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $SourceO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue
                $result =Import-PSSession -Session $SourceO365Session -AllowClobber -ErrorAction Stop -WarningAction silentlyContinue -DisableNameChecking -Prefix SRC 
                $msg = "SUCCESS: Connection to source Office 365 Remote PowerShell."
                Write-Host -ForegroundColor Green  $msg
                Log-Write -Message $msg -LogFile $logFile
            }
            until (($loginAttempts -ge 3) -or ($($SourceO365Session.State) -eq "Opened"))

            # Only 3 attempts allowed
            if($loginAttempts -ge 3) {
                $msg = "ERROR: Failed to connect to the Source Office 365. Review your source Office 365 admin credentials and try again."
                Write-Host $msg -ForegroundColor Red
                Log-Write -Message $msg -LogFile $logFile
                Start-Sleep -Seconds 5
                Exit
            }
        }
        catch {
            $msg = "ERROR: Failed to connect to source Office 365."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
            Get-PSSession | Remove-PSSession
            Exit
        }
        return $SourceO365Session
    } 
    else {
        Get-PSSession | Remove-PSSession
    }
}

# Function to create destination EXO PowerShell session
Function Connect-DestinationExchangeOnline {
    #Prompt for destination Office 365 global admin Credentials
    $msg = "INFO: Connecting to the destination Office 365 tenant."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile

    try {
        $loginAttempts = 0
        do {
            $loginAttempts++
            # Connect to Source Exchange Online
            if(!$O365Credentials) {
                $destinationO365Creds = Get-Credential -Message "Enter Your Source Office 365 Admin Credentials."
                if (!($destinationO365Creds)) {
                    $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg -LogFile $logFile
                    Exit
                }
            }
            else {
                $destinationO365Creds = Get-Credential $O365Credentials
            }
            $destinationO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $destinationO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue
            $result =Import-PSSession -Session $destinationO365Session -AllowClobber -ErrorAction Stop -WarningAction silentlyContinue -DisableNameChecking #-Prefix DST 
            $msg = "SUCCESS: Connection to destination Office 365 Remote PowerShell."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg -LogFile $logFile
        }
        until (($loginAttempts -ge 3) -or ($($destinationO365Session.State) -eq "Opened"))

        # Only 3 attempts allowed
        if($loginAttempts -ge 3) {
            $msg = "ERROR: Failed to connect to the destination Office 365. Review your destination Office 365 admin credentials and try again."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg -LogFile $logFile
            Start-Sleep -Seconds 5
            Exit
        }
    }
    catch {
        $msg = "ERROR: Failed to connect to destination Office 365."
        Write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Get-PSSession | Remove-PSSession
        Exit
    }
    return $destinationO365Session
}

# Function to export all Office 365 groups from the source tenant
 Function Get-O365groups {

    $msg = "INFO: Exporting Office 365 groups from source Office 365 tenant."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile

    #Export O365 groups from source O365 tenant. SRC prefix
    $exportO365Groups= @(Get-SRCUnifiedGroup | select Displayname,SharePointSiteUrl,primarysmtpaddress)

    $exportO365GroupsArray = @()

    Foreach($exportO365Group in $exportO365Groups) {

        $groupLineItem = New-Object PSObject
        $groupLineItem | Add-Member -MemberType NoteProperty -Name srcDisplayName -Value $exportO365Group.DisplayName
        $groupLineItem | Add-Member -MemberType NoteProperty -Name srcSharePointSiteUrl -Value $exportO365Group.SharePointSiteUrl
        $groupLineItem | Add-Member -MemberType NoteProperty -Name srcPrimarySmtpAddress -Value $exportO365Group.PrimarySmtpAddress
        $groupLineItem | Add-Member -MemberType NoteProperty -Name dstSharePointSiteUrl -Value ""
        $groupLineItem | Add-Member -MemberType NoteProperty -Name dstPrimarySmtpAddress -Value ""

        $exportO365GroupsArray += $groupLineItem
    }

    try {
        $exportO365GroupsArray| Export-Csv -Path $workingDir\ExportedO36Groups.csv -NoTypeInformation -force

        $msg = "SUCCESS: CSV file '$workingDir\ExportedO36Groups.csv' processed, exported and open."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg -LogFile $logFile
    }
    catch {
        $msg = "ERROR: Failed to export Office 365 groups to '$workingDir\ExportedO36Groups.csv' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    }

    $msg = "ACTION:  Please provide the 'dstSharePointSiteUrl' and 'dstPrimarySmtpAddress' in the opened CSV file and once you finish, save it."
    Write-Host -ForegroundColor Yellow  $msg
    Log-Write -Message $msg -LogFile $logFile

    try {
        #Open the CSV file for editing
        Start-Process -FilePath $workingDir\ExportedO36Groups.csv
    }
    catch {
        $msg = "ERROR: Failed to open '$workingDir\ExportedO36Groups.csv' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    }
    
    $msg = "ACTION:  Press any key to continue." 
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg -LogFile $logFile
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

    #Re-import the edited CSV file
    Try{
        $groups = @(Import-CSV "$workingDir\ExportedO36Groups.csv" | where-Object { $_.PSObject.Properties.Value -ne ""})
                 
        return $groups      
    }
    Catch [Exception] {
        $msg = "ERROR: Failed to import Office 365 groups from the CSV file '$workingDir\O365SendAsPermissions.csv'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg -LogFile $logFile
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    } 
 }

# Function to display the workgroups created by the user
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
            $workgroupsPage = @(Get-BT_Workgroup -PageOffset $workgroupOffSet -PageSize $workgroupPageSize -IsDeleted false) #-CreatedBySystemUserId $ticket.SystemUserId 
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC workgroups."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
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
                Log-Write -Message $msg -LogFile $logFile
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

# Function to display all customers
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
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
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
                Log-Write -Message $msg -LogFile $logFile
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

Function Get-CustomerUrlId {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$customerOrganizationId
    )

    $customerUrlId = (Get-BT_Customer -OrganizationId $customerOrganizationId).Id

    Return $customerUrlId

}

# Function to display all customers
Function Select-MSPC_EndUser {

    param 
    (      
        [parameter(Mandatory=$true)] [String]$customerOrganizationId
    )

    #######################################
    # Display all end users
    #######################################

    $endUserPageSize = 100
  	$endUserOffSet = 0
	$endUsers = $null

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC end users..."

    do
    {   
        try { 
            $endUsersPage = @(Get-BT_CustomerEndUser -Ticket $customerTicket -OrganizationID $customerOrganizationId -PageOffset $endUserOffSet -PageSize $endUserPageSize  -IsDeleted False -IsArchived False)
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC end users."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
            Exit
        }
    
        if($endUsersPage) {
            $endUsers += @($endUsersPage)
            foreach($endUser in $endUsersPage) {
                Write-Progress -Activity ("Retrieving end user (" + $endUsers.Length + ")") -Status $endUser.primaryAddress
            }

            $endUserOffSet += $endUserPageSize
        }

    } while($endUsersPage)

    if($endUsers -ne $null -and $endUsers.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $endUsers.Length.ToString() + " end user(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No end users found." 
        Exit
    }

    #######################################
    # {Prompt for the mailbox end user
    #######################################
    
    if($endUsers -ne $null)
    {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select an end user:" 

        for ($i=0; $i -lt $endUsers.Length; $i++)
        {
            $endUser = $endUsers[$i]
            #$endUser.primaryEmailAddress $endUser.PrimaryIdentity $endUser.EmailAddress,$endUser.UserPrincipalName
            Write-Host -Object $i,"-",$endUser.PrimaryIdentity
        }
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($endUsers.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0 or x")
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($endUsers.Length-1) + ", or x")
            }

            if($result -eq "x")
            {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $endUsers.Length))
            {
                $endUser=$endUsers[$result]
                Return $endUser.Id
            }
        }
        while($true)

    }

}

Function Display-MSPC_EndUsers {

    param 
    (      
        [parameter(Mandatory=$true)] [String]$customerOrganizationId
    )

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

    $customer = Get-BT_Customer -OrganizationId $customerOrganizationId -IsDeleted False -IsArchived False 
    $customerName = $customer.CompanyName

 
    Write-Host
    $msg = "INFO: Retrieving MSPC end users from '$customerName' customer..."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile

    $endUsers = @()
    $endUsersArray = @()

        # Retrieve all endUsers from the specified project
        $endUserOffSet = 0
        $endUserPageSize = 100
        $endUsers = $null

        do {
            $endUsersPage = @(Get-BT_CustomerEndUser -Ticket $customerTicket -OrganizationID $customerOrganizationId -PageOffset $endUserOffSet -PageSize $endUserPageSize  -IsDeleted False -IsArchived False)

            if($endUsersPage) {
                $endUsers += @($endUsersPage)

                foreach($endUser in $endUsersPage) {
                    Write-Progress -Activity ("Retrieving end users from '$customerName' customer") -Status $endUser.EmailAddress.ToLower()

                    $tab = [char]9
                    Write-Host -nonewline "Customer: "
                    Write-Host -nonewline -ForegroundColor Yellow  "$customerName  "               
                    write-host -nonewline "EmailAddress: "
                    write-host            -ForegroundColor Yellow "$($endUser.EmailAddress)"

                    $endUserLineItem = New-Object PSObject
                    $endUserLineItem | Add-Member -MemberType NoteProperty -Name CustomerName -Value $customerName
                    $endUserLineItem | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $endUser.EmailAddress
                    $endUserLineItem | Add-Member -MemberType NoteProperty -Name NewEmailAddress -Value ""

                    $endUsersArray += $endUserLineItem
                }

                $endUserOffSet += $endUserPageSize
            }
        } while($endUsersPage)

        if($endUsers -ne $null -and $endUsers.Length -ge 1) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $($endUsers.Length) end users found." 
        }
        else {
            Write-Host -ForegroundColor Red "INFO: No end users found. Script aborted." 
            Exit
        }

        try {
            $endUsersArray| Export-Csv -Path $workingDir\EndUsers.csv -NoTypeInformation -force

            $msg = "SUCCESS: CSV file '$workingDir\EndUsers.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg -LogFile $logFile
        }
        catch {
            $msg = "ERROR: Failed to export end users to '$workingDir\EndUsers.csv' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
            Exit
        }

        try {
            #Open the CSV file for editing
            Start-Process -FilePath $workingDir\EndUsers.csv
        }
        catch {
            $msg = "ERROR: Failed to open '$workingDir\EndUsers.csv' CSV file. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
            Exit
        }
    


}

# Function to display all endpoints under a customer
Function Select-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId,
        [parameter(Mandatory=$false)] [String]$endpointType,
        [parameter(Mandatory=$false)] [String]$endpointName,
        [parameter(Mandatory=$false)] [object]$endpointConfiguration,
        [parameter(Mandatory=$false)] [String]$exportOrImport,
        [parameter(Mandatory=$false)] [boolean]$deleteEndpointType

    )

    #####################################################################################################################
    # Display all MSPC endpoints. If $endpointType is provided, only endpoints of that type
    #####################################################################################################################

    $endpointPageSize = 100
  	$endpointOffSet = 0
	$endpoints = $null

    Write-Host
    if($endpointType -ne "") {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport $endpointType endpoints..."
    }else {
        Write-Host -Object  "INFO: Retrieving MSPC $exportOrImport endpoints..."
    }

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

    do {
        try{
            if($endpointType -ne "") {
                $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType)
            }else{
                $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize)
            }
        }
        catch {
            $msg = "ERROR: Failed to retrieve MSPC endpoints."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
            Exit
        }

        if($endpointsPage) {
            
            $endpoints += @($endpointsPage)

            foreach($endpoint in $endpointsPage) {
                Write-Progress -Activity ("Retrieving endpoint (" + $endpoints.Length + ")") -Status $endpoint.Name
            }
            
            $endpointOffset += $endpointPageSize
        }
    } while($endpointsPage)

    Write-Progress -Activity " " -Completed

    if($endpoints -ne $null -and $endpoints.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($endpoints.Length) endpoint(s) found."
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No endpoints found." 
    }

    #####################################################################################################################
    # Prompt for the endpoint. If no endpoints found and endpointType provided, ask for endpoint creation
    #####################################################################################################################
    if($endpoints -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the $exportOrImport $endpointType endpoint:" 

        for ($i=0; $i -lt $endpoints.Length; $i++) {
            $endpoint = $endpoints[$i]
            Write-Host -Object $i,"-",$endpoint.Name
        }
        Write-Host -Object "c - Create a new $endpointType endpoint"
        Write-Host -Object "x - Exit"
        Write-Host

        do
        {
            if($endpoints.count -eq 1) {
                $result = Read-Host -Prompt ("Select 0, c or x")
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($endpoints.Length-1) + ", c or x")
            }
            if($result -eq "c") {
                if ($endpointName -eq "") {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration                  
                }
                else {
                    $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
                }
                Return $endpointId
            }
            if($result -eq "x") {
                Exit
            }
            if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $endpoints.Length)) {
                $endpoint=$endpoints[$result]
                Return $endpoint.Id
            }
        }
        while($true)

    } 
    elseif($endpoints -eq $null -and $endpointType -ne "") {

        do {
            $confirm = (Read-Host -prompt "Do you want to create a $endpointType endpoint ?  [Y]es or [N]o")
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if($confirm.ToLower() -eq "y") {
            if ($endpointName -eq "") {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration 
            }
            else {
                $endpointId = create-MSPC_Endpoint -CustomerOrganizationId $CustomerOrganizationId -ExportOrImport $exportOrImport -ExportOrImport $exportOrImport -EndpointType $endpointType -EndpointConfiguration $endpointConfiguration -EndpointName $endpointName
            }
            Return $endpointId
        }
    }
}

# Function to get endpoint data
Function Get-MSPC_EndpointData {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId,
        [parameter(Mandatory=$true)] [guid]$endpointId
    )

    $customerTicket  = Get-BT_Ticket -OrganizationId $customerOrganizationId

    try {
        $endpoint = Get-BT_Endpoint -Ticket $customerTicket -Id $endpointId -IsDeleted False -IsArchived False  -ShouldUnmaskproperties $true | Select-Object -Property Name, Type -ExpandProperty Configuration
        
        $msg = "SUCCESS: Endpoint '$($endpoint.Name)' credentials retrieved." 
        write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg -LogFile $logFile 

        if($endpoint.Type -eq "AzureFileSystem") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername

            return $endpointCredentials        
        }
        
        elseif($endpoint.Type -eq "Pst") {

            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AccessKey -Value $endpoint.AccessKey
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name ContainerName -Value $endpoint.ContainerName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername

            return $endpointCredentials        
        }
        elseif($endpoint.Type -eq "OneDriveProAPI"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureStorageAccountName -Value $endpoint.AzureStorageAccountName
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AzureAccountKey -Value $endpoint.AzureAccountKey

            return $endpointCredentials   
        }
        elseif($endpoint.Type -eq "Office365Groups"){
            
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Url -Value $endpoint.Url
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            return $endpointCredentials     
        }
        elseif($endpoint.Type -eq "ExchangeOnline2"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword

            return $endpointCredentials  
        }
        elseif($endpoint.Type -eq "AzureSubscription"){
            $endpointCredentials = New-Object PSObject
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name Name -Value $endpoint.Name
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name UseAdministrativeCredentials -Value $endpoint.UseAdministrativeCredentials
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativeUsername -Value $endpoint.AdministrativeUsername
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name AdministrativePassword -Value $endpoint.AdministrativePassword
            $endpointCredentials | Add-Member -MemberType NoteProperty -Name SubscriptionID -Value $endpoint.SubscriptionID

            return $endpointCredentials
        
        }
        elseif($endpoint.Type -eq "DropBox"){
            $endpointCredentials = New-Object PSObject
            return $endpointCredentials
        }
    }
    catch {
        $msg = "ERROR: Failed to retrieve endpoint '$($endpoint.Name)' credentials."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
    }
 }
 
# Function to create an endpoint under a customer
# Configuration Table in https://www.bittitan.com/doc/powershell.html#PagePowerShellmspcmd%20
Function Create-MSPC_Endpoint {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$CustomerOrganizationId,
        [parameter(Mandatory=$false)] [String]$endpointType,
        [parameter(Mandatory=$false)] [String]$endpointName,
        [parameter(Mandatory=$false)] [object]$endpointConfiguration,
        [parameter(Mandatory=$false)] [String]$exportOrImport
    )

    $customerTicket  = Get-BT_Ticket -OrganizationId $customerOrganizationId

    if($endpointType -eq "AzureFileSystem"){
        
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")

            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
            }while ($secretKey -eq "")

            $msg = "INFO: Azure storage account access key is '$secretKey'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $azureAccountName; #Azure Storage Account Name        
                "AccessKey" = $secretKey; #Azure Storage Account SecretKey         
                "ContainerName" = $ContainerName #Container Name
            }
        }
        else {
            $azureFileSystemConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername; #Azure Storage Account Name        
                "AccessKey" = $endpointConfiguration.AccessKey; #Azure Storage Account SecretKey         
                "ContainerName" = $endpointConfiguration.ContainerName #Container Name
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureFileSystemConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
        }    
    }
    elseif($endpointType -eq "AzureSubscription"){
           
        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($secretKey -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            do {
                $azureSubscriptionID = (Read-Host -prompt "Please enter the Azure subscription ID").trim()
            }while ($azureSubscriptionID -eq "")

            $msg = "INFO: Azure subscription ID is '$azureSubscriptionID'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $adminUsername;     
                "AdministrativePassword" = $adminPassword;         
                "SubscriptionID" = $azureSubscriptionID
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureSubscriptionConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;  
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword;    
                "SubscriptionID" = $endpointConfiguration.SubscriptionID 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
        }   
    }
    elseif($endpointType -eq "Pst"){

        #####################################################################################################################
        # Prompt for endpoint data or retrieve it from $endpointConfiguration
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

             do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")

            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key ").trim()
            }while ($secretKey -eq "")


            do {
                $containerName = (Read-Host -prompt "Please enter the container name").trim()
            }while ($containerName -eq "")

            $msg = "INFO: Azure subscription ID is '$containerName'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $azureAccountName;     
                "AccessKey" = $secretKey;  
                "ContainerName" = $containerName;       
            }
        }
        else {
            $azureSubscriptionConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.AzureConfiguration' -Property @{  
                "UseAdministrativeCredentials" = $true;         
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;  
                "AccessKey" = $endpointConfiguration.AccessKey;    
                "ContainerName" = $endpointConfiguration.ContainerName 
            }
        }
        try {
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $customerTicket -Name $endpointName -Type $endpointType -Configuration $azureSubscriptionConfiguration 

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $checkEndpoint.Id
            }
            

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
        }  
    }
    elseif($endpointType -eq "OneDriveProAPI"){

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($secretKey -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            do {
                $azureAccountName = (Read-Host -prompt "Please enter the Azure storage account name").trim()
            }while ($azureAccountName -eq "")
        
            $msg = "INFO: Azure storage account name is '$azureAccountName'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            do {
                $secretKey = (Read-Host -prompt "Please enter the Azure storage account access key").trim()
            }while ($secretKey -eq "")
        
            $msg = "INFO: Azure storage account access key is '$secretKey'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile
    
            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword;
                "AzureStorageAccountName" = $azureAccountName;
                "AzureAccountKey" = $secretKey
            }
        }
        else {
            $oneDriveProAPIConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.SharePointOnlineConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword;
                "AzureStorageAccountName" = $endpointConfiguration.AzureStorageAccountName;
                "AzureAccountKey" = $endpointConfiguration.AzureAccountKey
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $oneDriveProAPIConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile               
        }
    }
    elseif($endpointType -eq "Office365Groups") {

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            do {
                $url = (Read-Host -prompt "Please enter the Office 365 group URL").trim()
            }while ($url -eq "")
        
            $msg = "INFO: Office 365 group URL is '$url'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile
        
        
            do {
                $adminUsername = (Read-Host -prompt "Please enter the admin email address").trim()
            }while ($adminUsername -eq "")
        
            $msg = "INFO: Admin email address is '$adminUsername'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            do {
                $adminPassword = (Read-Host -prompt "Please enter the admin password").trim()
            }while ($secretKey -eq "")
        
            $msg = "INFO: Admin password is '$adminPassword'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $office365GroupsConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.SharePointConfiguration" -Property @{
                "Url" = $url;
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $adminUsername;
                "AdministrativePassword" = $adminPassword
            }
        }
        else {
            $office365GroupsConfiguration = New-Object -TypeName "ManagementProxy.ManagementService.SharePointConfiguration" -Property @{
                "Url" = $endpointConfiguration.Url;
                "UseAdministrativeCredentials" = $true;
                "AdministrativeUsername" = $endpointConfiguration.AdministrativeUsername;
                "AdministrativePassword" = $endpointConfiguration.AdministrativePassword
            }
        }

        try {
            
            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $office365GroupsConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $checkEndpoint.Id
            }
        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile               
        }
    }
    elseif($endpointType -eq "DropBox"){

        #####################################################################################################################
        # Prompt for endpoint data. 
        #####################################################################################################################
        if($endpointConfiguration -eq $null) {
            if ($endpointName -eq "") {
                do {
                    $endpointName = (Read-Host -prompt "Please enter the $exportOrImport endpoint name").trim()
                }while ($endpointName -eq "")
            }

            #####################################################################################################################
            # Create endpoint. 
            #####################################################################################################################

            $dropBoxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.DropBoxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativePassword" = ""
            }
        }
        else {
            $dropBoxConfiguration = New-Object -TypeName 'ManagementProxy.ManagementService.DropBoxConfiguration' -Property @{              
                "UseAdministrativeCredentials" = $true;
                "AdministrativePassword" = ""
            }
        }

        try {

            $checkEndpoint = Get-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -IsDeleted False -IsArchived False 

            if( $($checkEndpoint.Count -eq 0)) {

                $endpoint = Add-BT_Endpoint -Ticket $CustomerTicket -Name $endpointName -Type $endpointType -Configuration $dropBoxConfiguration

                $msg = "SUCCESS: The $exportOrImport $endpointType endpoint '$endpointName' created."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $endpoint.Id
            }
            else {
                $msg = "WARNING: $endpointType endpoint '$endpointName' already exists. Skipping endpoint creation."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg -LogFile $logFile

                Return $checkEndpoint.Id
            }

        }
        catch {
            $msg = "         ERROR: Failed to create the $exportOrImport $endpointType endpoint '$endpointName'."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile               
        }
    }
}

# Function to create a connector under a customer
Function Create-MW_Connector {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$CustomerOrganizationId,
        [parameter(Mandatory=$true)] [String]$ProjectName,
        [parameter(Mandatory=$true)] [String]$ProjectType,
        [parameter(Mandatory=$true)] [String]$importType,
        [parameter(Mandatory=$true)] [String]$exportType,   
        [parameter(Mandatory=$true)] [guid]$exportEndpointId,
        [parameter(Mandatory=$true)] [guid]$importEndpointId,  
        [parameter(Mandatory=$true)] [object]$exportConfiguration,
        [parameter(Mandatory=$true)] [object]$importConfiguration,
        [parameter(Mandatory=$false)] [String]$advancedOptions,   
        [parameter(Mandatory=$false)] [String]$folderFilter,
        [parameter(Mandatory=$false)] [String]$maximumSimultaneousMigrations  
        
    )

    try { 
        $connector = Add-MW_MailboxConnector -ticket $global:mwTicket `
        -UserId $global:mwTicket.UserId `
        -OrganizationId $CustomerOrganizationId `
        -Name $ProjectName `
        -ProjectType $ProjectType `
        -ExportType $exportType `
        -ImportType $importType `
        -SelectedExportEndpointId $exportEndpointId `
        -SelectedImportEndpointId $importEndpointId `
        -ExportConfiguration $exportConfiguration `
        -ImportConfiguration $importConfiguration `
        -AdvancedOptions $advancedOptions `
        -FolderFilter $folderFilter `
        -MaximumDataTransferRate ([int]::MaxValue) `
        -MaximumDataTransferRateDuration 600000 `
        -MaximumSimultaneousMigrations $maximumSimultaneousMigrations `
        -PurgePeriod 180 `
        -MaximumItemFailures 100 

        $msg = "SUCCESS: Connector '$($connector.Name)' created." 
        write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg -LogFile $logFile

        return $connector.Id
    }
    catch{
        $msg = "ERROR: Failed to create mailbox connector '$($connector.Name)'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile 
    }

}

# Function to clone an existing connector under a customer
function Clone-MW_Project {

    param 
    (      
        [parameter(Mandatory=$true)] [MigrationProxy.WebApi.Entity]$projectToClone
    )

    $newId = [guid]::NewGuid()
 
    $projectToClone.Name = $projectToClone.Name + "_CLONED"
    $selectedProject_ImportConf = $projectToClone | select -ExpandProperty ImportConfiguration
    $selectedProject_ExportConf = $projectToClone | select -ExpandProperty ExportConfiguration

    $newJob = Add-MW_MailboxConnector -ticket $global:mwTicket -Name $projectToClone.Name -ProjectType $projectToClone.projecttype `
    -ImportType $projectToClone.ImportType -ExportConfiguration $selectedProject_ExportConf `
    -ExportType $projectToClone.ExportType -ImportConfiguration $selectedProject_ImportConf `
    -SelectedExportEndpointId $projectToClone.SelectedExportEndpointId `
    -SelectedImportEndpointId $projectToClone.SelectedImportEndpointId `
    -OrganizationId $projectToClone.OrganizationId -UserId $projectToClone.UserId `
    -ZoneRequirement $projectToClone.ZoneRequirement -MaxLicensesToConsume $projectToClone.MaxLicensesToConsume `
    -AdvancedOptions $projectToClone.AdvancedOptions -MaximumItemFailures $projectToClone.MaximumItemFailures `
    -ErrorAction Stop

    $msg = "SUCCESS: Mailbox connector '$($projectToClone.Name)' created." 
    write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg -LogFile $logFile
 
    return $newJob
}
 
# Function to compare 2 existing connectors under the same customer
Function Compare_MW_connectors {
    param 
    (      
        [parameter(Mandatory=$true)] [MigrationProxy.WebApi.Entity]$sourceConnector,
        [parameter(Mandatory=$true)] [MigrationProxy.WebApi.Entity]$targetConnector

    )

    $sourceProjectType = $sourceConnector.projecttype
    $sourceImportType = $sourceConnector.ImportType
    $sourceExportType = $sourceConnector.ExportType
    $targetProjectType = $targetConnector.projecttype
    $targetImportType = $targetConnector.ImportType
    $targetExportType = $targetConnector.ExportType


    if(($sourceProjectType -eq $targetProjectType) -and ($sourceImportType -eq $targetImportType)  -and ($sourceExportType -eq $targetExportType)) {
        Return $true
    }
    else{
        $msg = "ERROR: Target connector type ($targetProjectType : $targetExportType->$targetImportType) does not match with source connector type ($sourceProjectType : $sourceExportType->$sourceImportType)."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Return False
    }


 
 }

Function Get-CsvFile {
    Write-Host
    Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import Office 365 groups (Press cancel to create one)"
    Get-FileName $workingDir

    # Import CSV and validate if headers are according the requirements
    try {
        $groups = Import-Csv $global:inputFile
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
    if ( $groups.count -eq 0 ) {
        $msg = "ERROR: '$global:inputFile' CSV file exist but it is empty. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Exit
    }

    # Validate CSV Headers
    $CSVHeaders = @("DisplayName","srcSharePointSiteUrl","srcPrimarySmtpAddress","dstSharePointSiteUrl","dstPrimarySmtpAddress")
    foreach ($header in $CSVHeaders) {
        if ($groups.$header -eq "" ) {
            $msg = "ERROR: '$global:inputFile' CSV file does not have all the required columns. Required columns are: '$($CSVHeaders -join "', '")'. Script aborted."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Exit
        }
    }
 
 }

# Function to get a CSV file name or to create a new CSV file
Function Get-FileName {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$initialDirectory

    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $global:inputFile = $OpenFileDialog.filename

    if($OpenFileDialog.filename -ne "") {		    
        $msg = "SUCCESS: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg -LogFile $logFile
    }
}

Function Menu-MigrationSubmission() {
    param 
    (      
        [parameter(Mandatory=$true)] [Array]$allMailboxes,
        [parameter(Mandatory=$true)] [Array]$projectName

    )

     # Select which mailboxes have to be submitted
    Write-Host
    Write-Host -ForegroundColor Yellow "ACTION: Which migrations would you like to submit:" 
    Write-Host "0 - All migrations"
    Write-Host "1 - Not started migrations"
    Write-Host "2 - Failed migrations"
    Write-Host "3 - Successful migrations that contain errors"
    if($ProjectName -match "FS-DropBox-") {
        Write-Host "4 - Specify the email address of DropBox account."
    }
    elseif($ProjectName -match "Mailbox-O365 Groups conversations") {
        Write-Host "4 - Specify the email address of the Office 365 Group."
    }
    elseif($ProjectName -match "FS-OD4B-") {
        Write-Host "4 - Specify the email address of OneDrive For Business account."
    }
    elseif($ProjectName -match "PST-O365-"){
        Write-Host "4 - Specify the email address of the Office 365 mailbox."
    }
    Write-Host "5 - All migrations that were not successful (failed or stopped)"
    Write-Host "x - Exit"
    Write-Host

    $continue=$true
    do {
        $result = Read-Host -Prompt "Select 0-5 or x"
        if($result -eq "x") {
            Exit
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 5)) {
            $statusAction = [int]$result
            $continue=$false
        }
    } while($continue)

    $count = 0
    $mailboxToSubmit = $null

    # Select migration pass type
    Write-Host
    Write-Host  -ForegroundColor Yellow "ACTION: What type of migration would you like to perform:"
    Write-Host "0 - Migration (including delta pass if previously migrated)"
    Write-Host "1 - Verify credentials"
    Write-Host "2 - Retry errors"
    Write-Host "3 - Trial migration"
    Write-Host "x - Exit"
    Write-Host

    $continue=$true
    do {
        $result = Read-Host -Prompt "Select 0-3 or x" 
        if($result -eq "x") {
            return $null
        }
        if(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -le 3)) {
            switch([int]$result) {
                0
                {
                    $migrationAction = [MigrationProxy.WebApi.MailboxQueueTypes]::Full
                    $continue=$false
                }

                1
                {
                    $migrationAction = [MigrationProxy.WebApi.MailboxQueueTypes]::Verification
                    $continue=$false
                }

                2
                {
                    $migrationAction = [MigrationProxy.WebApi.MailboxQueueTypes]::Repair
                    $continue=$false
                }

                3
                {
                    $migrationAction = [MigrationProxy.WebApi.MailboxQueueTypes]::Trial
                    $continue=$false
                }
            }

        }
    } while ($continue)

    $migrationType = $migrationAction[0]

    # If only one mailbox has to be submitted
    if($statusAction -eq 4)
    {
        if($ProjectName -match "FS-DropBox-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the DropBox account to submit:  "  -NoNewline
        }
        elseif($ProjectName -match "Mailbox-O365 Groups conversations") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the Office 365 Group to submit:  "  -NoNewline
        }
        elseif($ProjectName -match "FS-OD4B-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the OneDrive For Business account to submit:  "  -NoNewline
        }
        elseif($ProjectName -match "PST-O365-") {
            Write-Host -ForegroundColor Yellow "ACTION: Email address of the PST file to submit:  "  -NoNewline
        }
        
        $emailAddress = Read-Host

        if($emailAddress.Length -ge 1) {
            $mailboxToSubmit = $emailAddress
            $msg = "INFO: The specified email address is '$emailAddress'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile $msg
        }

        if($mailboxToSubmit -eq $null -or $mailboxToSubmit.Length -eq 0) {
            Write-Host "ERROR: No email address was entered" -ForegroundColor Red
            Return
        }

    }

    # Submitting mailboxes for migration    
    Write-Host
    Write-Host "INFO: Submitting migrations..."

    $count=0
    $submittedcount=0
    foreach($mailbox in $allMailboxes)
    {
        $submit = $false
        $status = "NotMigrated"

        $count++
        if(($ProjectName -match "FS-DropBox-" -or $ProjectName -match "FS-OD4B-" -or $ProjectName -match "PST-O365-")-and $mailbox.ImportEmailAddress -ne "") {
            Write-Progress -Activity ("Submitting migrations (" + $count + "/" + $allMailboxes.Length + ")") -Status $mailbox.ImportEmailAddress.ToLower() -PercentComplete ($count/$allMailboxes.Length*100)
        }

        if($ProjectName -match "Mailbox-O365 Groups conversations" -and $mailbox.ExportEmailAddress -ne "") {
            Write-Progress -Activity ("Submitting migrations (" + $count + "/" + $allMailboxes.Length + ")") -Status $mailbox.ExportEmailAddress.ToLower() -PercentComplete ($count/$allMailboxes.Length*100)
        }
        elseif($ProjectName -match "Document-" -and $mailbox.ExportLibrary -ne "") {
            Write-Progress -Activity ("Submitting migrations (" + $count + "/" + $allMailboxes.Length + ")") -Status $mailbox.ExportLibrary.ToLower() -PercentComplete ($count/$allMailboxes.Length*100)
        }


        # Get the latest status of each of the migrations
        if($statusAction -ne 4)
        {   
            try { 
                $latestMigration = Get-MW_MailboxMigration -Ticket $mwTicket -MailboxId $mailbox.Id -PageSize 1 -PageOffset 0
            }
            catch {
                $msg = "ERROR: Failed to retrieve the latest status of each of the migrations."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg -LogFile $logFile $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message -LogFile $logFile
                Exit
            }
            if($latestMigration -ne $null)
            {
                $status = $latestMigration.Status
            }
        }
        # Get the latest status of the specified email address
        elseif($statusAction -eq 4)
        {   
             
            if($mailboxToSubmit -eq $mailbox.ExportEmailAddress -or $mailboxToSubmit -eq $mailbox.ImportEmailAddress) {
                try { 
                    $latestMigration = Get-MW_MailboxMigration -Ticket $mwTicket -MailboxId $mailbox.Id -PageSize 1 -PageOffset 0
                }
                catch {
                    $msg = "ERROR: Failed to retrieve the latest status of '$mailboxToSubmit'."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg -LogFile $logFile $msg
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $_.Exception.Message -LogFile $logFile
                    Exit
                }
                if($latestMigration -ne $null)
                {
                    $status = $latestMigration.Status
                }
            }
        }


        switch($status)
        {
            "NotMigrated"
            {
                if($statusAction -eq 0 -or $statusAction -eq 1 -or $statusAction -eq 5) 
                {
                    $submit = $true
                }
                elseif($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1)
                {
                    if(($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()))
                    {
                        $submit = $true
                    }
                }
            }

            "Completed"
            {
                if($statusAction -eq 0)
                {
                    $submit = $true
                }
                # Only successfully completed migrations with errors
                elseif($statusAction -eq 3)
                {
                    $stats = MW-GetMailboxStats -mailbox $mailbox

                    $Calendar = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Calendar)
                    $Contact = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Contact)
                    $Mail = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Mail)
                    $Journal = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Journal)
                    $Note = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Note)
                    $Task = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Task)
                    $Folder = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Folder)
                    $Rule = [int]([MigrationProxy.WebApi.MailboxItemTypes]::Rule)

                    if($stats -ne $null) {
                        foreach($info in $stats.MigrationStatsInfos) {
                            switch ([int]$info.ItemType) {
                                $Folder {
                                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $folderErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Calendar {
                                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $calendarErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Contact {
                                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $contactErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Mail {
                                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $mailErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Task {
                                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $taskErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Note {
                                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $noteErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Journal {
                                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $journalErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                $Rule {
                                    if($info.TaskType -eq [MigrationProxy.WebApi.TaskType]::Import) {
                                        $ruleErrorCount = $info.MigrationStats.ErrorCount + $info.MigrationStats.ErrorCountTotal
                                    }
                                    break
                                }

                                default {break}
                            }
                        }


                        $totalErrorCount = $folderErrorCount + $calendarErrorCount + $contactErrorCount + $mailErrorCount + $taskErrorCount + $noteErrorCount + $journalErrorCount + $rulesErrorCount
                    }

                    if($totalErrorCount -ge 1)
                    {
                        $submit = $true
                    }
                }
                elseif($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1)
                {
                    if(($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()))
                    {
                        $submit = $true
                    }
                }
            }

            "Failed"
            {
                if($statusAction -eq 0 -or $statusAction -eq 2 -or $statusAction -eq 5)
                {
                    $submit = $true
                }
                elseif($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1)
                {
                    if(($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()))
                    {
                        $submit = $true
                    }
                }
            }

            "Stopped"
            {
                if($statusAction -eq 0 -or $statusAction -eq 5)
                {
                    $submit = $true
                }
                elseif($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1)
                {
                    if(($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()))
                    {
                        $submit = $true
                    }
                }
            }

            "MaximumTransferReached"
            {
                if($statusAction -eq 0 -or $statusAction -eq 5)
                {
                    $submit = $true
                }
                elseif($statusAction -eq 4 -and $mailboxToSubmit -ne $null -and $mailboxToSubmit.Length -ge 1)
                {
                    if(($mailbox.ExportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()) -or ($mailbox.ImportEmailAddress.ToLower() -eq $mailboxToSubmit.ToLower()))
                    {
                        $submit = $true
                    }
                }
            }
        }

        if($submit) { 
               $migration = Add-MW_MailboxMigration -Ticket $mwTicket -MailboxId $mailbox.Id -ConnectorId $mailbox.ConnectorId -Type $migrationType -UserId $mwTicket.UserId -Priority 1 -Status Submitted -errorAction SilentlyContinue   
                           
               #If error has occurred 
               if(!$migration){
                   if($mailbox.ExportEmailAddress -ne "") {
                       $msg = "ERROR: Failed to submit migration '$($mailbox.ExportEmailAddress)' because it is currently running."
                       Write-Host -ForegroundColor Red  $msg
                       Log-Write -Message $msg -LogFile $logFile $msg
                   }
                   elseif($mailbox.ExportLibrary -ne "") {
                       $connector = Get-MW_MailboxConnector -Ticket $mwTicket -Id $mailbox.ConnectorId 
                       $msg = "ERROR: Failed to submit migration '$($mailbox.ExportLibrary)' in '$($connector.Name)' because it is currently running."
                       Write-Host -ForegroundColor Red  $msg
                       Log-Write -Message $msg -LogFile $logFile $msg
                   }             
               }else{
                    $submittedcount += 1
               }
        }
    } 

    Write-Host  -ForegroundColor Green "SUCCESS: $submittedcount out of $count migrations were submitted for migration"

    return $count 
 }
 
# Function to query destination email addresses
Function Query-EmailAddressMapping {
    do {
        $confirm = (Read-Host -prompt "Are you migrating to the same email addresses?  [Y]es or [N]o")
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($confirm.ToLower() -eq "y") {
        $global:sameEmailAddresses = $true
        $global:sameUserName = $true
        $global:differentDomain = $false
        
        $msg = "WARNING: Since you are migrating the domain to the destination Office 365 tenant,`r`n         either the source or destination primary email addresses must be in onmicrosoft.com format."
        Write-Host -ForegroundColor Yellow $msg      
        
        $script:destinationDomain = (Read-Host -prompt "Please enter the current destination domain")
        $msg = "INFO: Current destination domain is '$script:destinationDomain'."
        Write-Host $msg
        Log-Write -Message $msg -LogFile $logFile
          
    }
    elseif($confirm.ToLower() -eq "n") {
        
        $global:sameEmailAddresses = $false

        do {
            $confirm = (Read-Host -prompt "Are you migrating to a different domain?  [Y]es or [N]o")
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        # If destination Domain is different
        if($confirm.ToLower() -eq "y") {
            $global:differentDomain = $true
            if($createDistributionGroups) {
                do {
                    $script:destinationDomain = (Read-Host -prompt "Please enter the destination domain")
                }while ($script:destinationDomain -eq "")
                $msg = "INFO: Destination domain is '$script:destinationDomain'."
                Write-Host $msg
                Log-Write -Message $msg -LogFile $logFile
            }
            else {
                do{
                    $confirm = (Read-Host -prompt "Are the destination email addresses keeping the same user prefix?  [Y]es or [N]o")
                } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

                if($confirm.ToLower() -eq "y") {
                    $global:sameUserName = $true

                    do {
                        $script:destinationDomain = (Read-Host -prompt "Please enter the destination domain")
                    }while ($script:destinationDomain -eq "")
                     $msg = "INFO: Destination domain is '$script:destinationDomain'."
                     Write-Host $msg
                     Log-Write -Message $msg -LogFile $logFile
                }
                else {
                    $global:sameUserName = $false

                    $msg = "WARNING: Since you are migrating to a different domain`r`n         but you are not keeping the user prefixes the same,`r`n         you will have to manually provide the current destination email addresses."
                    Write-Host -ForegroundColor Yellow $msg     
                    Log-Write -Message $msg -LogFile $logFile
                }    
            }        
        } 
        # If destination domain is the same but user prefix is different, source and destination email addresses must be in onmicrosoft.com format
        else {
            $global:differentDomain = $false
            $global:sameUserName = $false

            
            $msg = "WARNING: Since you are migrating the domain to the destination Office 365 tenant`r`n         but you are not keeping the user prefixes the same,`r`n         you will have to manually provide the current destination email addresses."
            Write-Host -ForegroundColor Yellow $msg     
            Log-Write -Message $msg -LogFile $logFile
        }   
    }
}

# Function to get the licenses of each of the Office 365 users
Function Get-OD4BAccounts {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    try {
        #Prompt for destination Office 365 global admin Credentials
        $msg = "INFO: Connecting to Azure Active Directory."
        Write-Host $msg
        Log-Write -Message $msg -LogFile $logFile

	    Connect-MsolService -Credential $Credentials -ErrorAction Stop
	    
        $msg = "SUCCESS: Connection to Azure Active Directory."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg -LogFile $logFile
    }
    catch {
	    $msg = "ERROR: Failed to connect to Azure Active Directory."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg -LogFile $logFile
        Start-Sleep -Seconds 5
        Exit
	}

    Write-Host
    $msg = "INFO: Exporting all users with OneDrive For Business."
    write-Host $msg
    Log-Write -Message $msg -LogFile $logFile

    $od4bArray = @()

    $allUsers = Get-MsolUser -All | Select-Object UserPrincipalName,primarySmtpAddress -ExpandProperty Licenses
    ForEach ($user in $allUsers) {
        $userUpn = $user.UserPrincipalName
        $accountSkuId = $user.AccountSkuId
        $services = $user.ServiceStatus

        ForEach ($service in $services) {
            $serviceType = $service.ServicePlan.ServiceType
            $provisioningStatus = $service.ProvisioningStatus

                if($serviceType -eq "SharePoint" -and $provisioningStatus -ne "disabled") {
                    $properties = @{UserPrincipalName=$userUpn
                                    Office365Plan=$accountSkuId
                                    Office365Service=$serviceType
                                    ServiceStatus=$provisioningStatus
                                    SourceFolder=$userUpn.split("@")[0]
                                    }

                    $obj1 = New-Object –TypeName PSObject –Property $properties 

                    $od4bArray += $obj1 
                    Break
                }            
        }
    }

    Return $od4bArray 
}

# Function to query destination email addresses
Function Query-O365GroupEmailAddressMapping {
    do {
        $confirm = (Read-Host -prompt "Are you migrating to the same email addresses?  [Y]es or [N]o")
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($confirm.ToLower() -eq "y") {
        $global:sameEmailAddresses = $true
        $global:sameUserName = $true
        $global:differentDomain = $false
        
        do {
            $domains = (Read-Host -prompt "Please enter the destination vanity domain (or domains separated by comma)")

            $global:destinationDomains = @($domains.split(","))

        }while ($global:destinationDomains -eq "")
        
        $msg = "INFO: Destination domain is '$global:destinationDomains'."
        Write-Host $msg
        Log-Write -Message $msg -LogFile $logFile

    }
    elseif($confirm.ToLower() -eq "n") {
        
        $global:sameEmailAddresses = $false

        do {
            $confirm = (Read-Host -prompt "Are you migrating to a different domain?  [Y]es or [N]o")
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        # If destination Domain is different
        if($confirm.ToLower() -eq "y") {
            
            $global:differentDomain = $true

            do {
                $domains = (Read-Host -prompt "Please enter the destination vanity domain (or domains separated by comma)")

                $global:destinationDomains = @($domains.split(","))

            }while ($global:destinationDomains -eq "")
            
            $msg = "INFO: Destination domain is '$global:destinationDomains'."
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile


            do{
                $confirm = (Read-Host -prompt "Are the destination email addresses keeping the same user prefix?  [Y]es or [N]o")
            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

            if($confirm.ToLower() -eq "y") {
                $global:sameUserName = $true
            }
            else {
                $global:sameUserName = $false
            } 
        }
        else {

            $global:differentDomain = $false
            $global:sameUserName = $false            
        }        
    }
}

# Function to export all Office 365 groups from the source tenant
Function Get-O365groups {

    query-O365GroupEmailAddressMapping

    write-host

    $msg = "INFO: Exporting Office 365 groups from source Office 365 tenant."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile

    #Export O365 groups from source O365 tenant. SRC prefix
    $exportO365Groups= @(Get-SRCUnifiedGroup -ResultSize Unlimited | select Displayname,SharePointSiteUrl,primarysmtpaddress)

    $exportO365GroupsArray = @()

    Foreach($exportO365Group in $exportO365Groups) {

        $groupLineItem = New-Object PSObject
        $groupLineItem | Add-Member -MemberType NoteProperty -Name srcDisplayName -Value $exportO365Group.DisplayName
        $groupLineItem | Add-Member -MemberType NoteProperty -Name srcSharePointSiteUrl -Value $exportO365Group.SharePointSiteUrl
        $groupLineItem | Add-Member -MemberType NoteProperty -Name srcPrimarySmtpAddress -Value $exportO365Group.PrimarySmtpAddress
        
        if($global:sameEmailAddresses -eq $true -and $exportO365Group.PrimarySmtpAddress -notmatch ".onmicrosoft.com" -and $global:destinationDomains.count -eq 1) {
            $groupLineItem | Add-Member -MemberType NoteProperty -Name dstSharePointSiteUrl -Value ""
            $groupLineItem | Add-Member -MemberType NoteProperty -Name dstPrimarySmtpAddress -Value $exportO365Group.PrimarySmtpAddress
        }
        elseif($global:sameEmailAddresses -eq $true -and $exportO365Group.PrimarySmtpAddress -match ".onmicrosoft.com" -and $global:destinationDomains.count -eq 1) {

            $exportO365GroupPrefix = $exportO365Group.PrimarySmtpAddress.split("@")[0]
            $destinationDomain = $global:destinationDomains

            $groupLineItem | Add-Member -MemberType NoteProperty -Name dstSharePointSiteUrl -Value ""
            $groupLineItem | Add-Member -MemberType NoteProperty -Name dstPrimarySmtpAddress -Value "$exportO365GroupPrefix@$destinationDomain"
        }
        elseif(!$global:sameEmailAddresses -eq $true -and $global:sameUserName -eq $true -and $global:destinationDomains.count -eq 1) {
            $groupLineItem | Add-Member -MemberType NoteProperty -Name dstSharePointSiteUrl -Value ""

            $exportO365GroupPrefix = $exportO365Group.PrimarySmtpAddress.split("@")[0]
            $destinationDomain = $global:destinationDomains

            $groupLineItem | Add-Member -MemberType NoteProperty -Name dstPrimarySmtpAddress -Value "$exportO365GroupPrefix@$destinationDomain"
        }
        else {
            $groupLineItem | Add-Member -MemberType NoteProperty -Name dstSharePointSiteUrl -Value ""
            $groupLineItem | Add-Member -MemberType NoteProperty -Name dstPrimarySmtpAddress -Value ""
        }

        $exportO365GroupsArray += $groupLineItem
    }

    try {
        $exportO365GroupsArray| Export-Csv -Path $workingDir\ExportedO36Groups.csv -NoTypeInformation -force

        $msg = "SUCCESS: CSV file '$workingDir\ExportedO36Groups.csv' processed, exported and open."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg -LogFile $logFile
    }
    catch {
        $msg = "ERROR: Failed to export Office 365 groups to '$workingDir\ExportedO36Groups.csv' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    }

    $msg = "ACTION:  Please provide the 'dstSharePointSiteUrl' and 'dstPrimarySmtpAddress' in the opened CSV file and once you finish, save it."
    Write-Host -ForegroundColor Yellow  $msg
    Log-Write -Message $msg -LogFile $logFile

    try {
        #Open the CSV file for editing
        Start-Process -FilePath $workingDir\ExportedO36Groups.csv
    }
    catch {
        $msg = "ERROR: Failed to open '$workingDir\ExportedO36Groups.csv' CSV file. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    }
    
    $msg = "ACTION:  Press any key to continue." 
    Write-Host -ForegroundColor Yellow $msg
    Log-Write -Message $msg -LogFile $logFile
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

    #Re-import the edited CSV file
    Try{
        $groups = @(Import-CSV "$workingDir\ExportedO36Groups.csv" | where-Object { $_.PSObject.Properties.Value -ne ""})
                 
        return $groups      
    }
    Catch [Exception] {
        $msg = "ERROR: Failed to import Office 365 groups from the CSV file '$workingDir\ExportedO36Groups.csv'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg -LogFile $logFile
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    } 
 }

Function Get-TenantDomain {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    try {
	    Connect-MsolService -Credential $Credentials -ErrorAction Stop
	    
        $tenantDomain = (Get-MsolDomain |?{$_.Name -match 'onmicrosoft.com'}).Name

    }
    catch {
	    $msg = "ERROR: Failed to connect to Azure Active Directory."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg -LogFile $logFile
        Start-Sleep -Seconds 5
        Exit
	}

    Return $tenantDomain
}

Function Get-VanityDomains {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    try {
	    Connect-MsolService -Credential $Credentials -ErrorAction Stop
	    
        $tenantDomains = @(Get-MsolDomain |?{$_.Name -notmatch 'onmicrosoft.com'}).Name
    }
    catch {
	    $msg = "ERROR: Failed to connect to Azure Active Directory."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg -LogFile $logFile
        Start-Sleep -Seconds 5
        Exit
	}

    Return $tenantDomains
}

# Function to wait for the user to press any key to continue
Function WaitForKeyPress{
    param 
    (      
        [parameter(Mandatory=$true)] [string]$message
    )
    
    Write-Host $message
    Log-Write -Message $message -LogFile $logFile
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

# Function to unzip a file
Function Unzip-File {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$zipfile
    )

    $folderName = (Get-Item $zipfile).Basename
    $fileName = $($zipfile.split("\")[-1])

    $result = New-Item -ItemType directory -Path $folderName -Force 

    try {
        $result = Expand-Archive $zipfile -DestinationPath $folderName -Force

        $msg = "SUCCESS: '$fileName' file unzipped into '$PSScriptRoot\$folderName'."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg -LogFile $logFile
    }
    catch {
        $msg = "ERROR: Failed to unzip '$fileName' file."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    }
}

# Function to download a file from a URL
Function Download-File {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$url,
        [parameter(Mandatory=$true)] [String]$outFile
    )

    $fileName = $($url.split("/")[-1])
    $folderName = $fileName.split(".")[0]

    $msg = "INFO: Downloading the latest version of '$fileName' agent (~12MB) from BitTitan..."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile

    #Download the latest version of UploaderWiz from BitTitan server
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    try {
        $result = Invoke-WebRequest -Uri $url -OutFile $outFile
        $msg = "SUCCESS: '$fileName' file downloaded into '$PSScriptRoot'."
        Write-Host -ForegroundColor Green $msg
        Log-Write -Message $msg -LogFile $logFile
    }
    catch {
        $msg = "ERROR: Failed to download '$fileName'."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile   
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Unzip-File $outFile 

    #Open the zip file 
    try {
    
            Start-Process -FilePath "$PSScriptRoot\$folderName"


        }
        catch {
            $msg = "ERROR: Failed to open '$PSScriptRoot' folder."
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg -LogFile $logFile
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message -LogFile $logFile
            Exit
        }

    
   # else {
   #     $msg = 
   #     "ERROR: Failed to download  UploaderWiz agent from BitTitan."
   #     Write-Host -ForegroundColor Red  $msg
   #     Log-Write -Message $msg -LogFile $logFile
   # }
 
 }

# Function to check if AzureRM is installed
Function Check-AzureRM {
     Try {
        $result = get-module -ListAvailable -name AzureRM -ErrorAction Stop
        if ($result) {
            $msg = "INFO: Ready to execute Azure PowerShell module $($result.moduletype), $($result.version), $($result.name)"
            Write-Host $msg
            Log-Write -Message $msg -LogFile $logFile
        }
        Else {
            $msg = "INFO: AzureRM module is not installed."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg -LogFile $logFile

            Install-Module AzureRM
            Import-Module AzureRM

            Try {
                
                $result = get-module -ListAvailable -name AzureRM -ErrorAction Stop
                
                If ($result) {
                    write-information "INFO: Ready to execute PowerShell module $($result.moduletype), $($result.version), $($result.name)"
                }
                Else {
                    $msg = "ERROR: Failed to install and import the AzureRM module. Script aborted."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg -LogFile $logFile   
                    Write-Host -ForegroundColor Red $_.Exception.Message
                    Log-Write -Message $_.Exception.Message -LogFile $logFile
                    Exit
                }
            }
            Catch {
                $msg = "ERROR: Failed to check if the AzureRM module is installed. Script aborted."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg -LogFile $logFile   
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message -LogFile $logFile
                Exit
            }
        }

    }
    Catch {
        $msg = "ERROR: Failed to check if the AzureRM module is installed. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    } 
}

# Function to connect to Azure
Function Connect-Azure{
    param(
        [Parameter(Mandatory=$true)] [PSObject]$azureCredentials,
        [Parameter(Mandatory=$false)] [String]$subscriptionID
    )

    $msg = "INFO: Connecting to Azure to create a blob container."
    Write-Host $msg
    Log-Write -Message $msg -LogFile $logFile
    Try {
        if($subscriptionID -eq $null) {
            $result = Login-AzureRMAccount -Credential $azureCredentials -Environment "AzureCloud" -ErrorAction Stop
        }
        else {
            $result = Login-AzureRMAccount -Credential $azureCredentials -Environment "AzureCloud" -SubscriptionId $subscriptionID -ErrorAction Stop
        }

        $azureAccount = (Get-AzureRmContext).Account
        $subscriptionName = (Get-AzureRmSubscription -SubscriptionID $subscriptionID).Name
        $msg = "SUCCESS: Connection to Azure: Account: $azureAccount Subscription: '$subscriptionName'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg -LogFile $logFile
    }
    catch {
        $msg = "ERROR: Failed to connect to Azure. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit
    }

}

# Function to check if a StorageAccount exists
Function Check-StorageAccount{
    param 
    (      
        [parameter(Mandatory=$true)] [String]$storageAccountName
    )   

    try {
        $storageAccount = Get-AzureRmStorageAccount -ErrorAction Stop |? {$_.StorageAccountName -eq $storageAccountName}

        $msg = "SUCCESS: Azure storage account '$storageAccountName' found."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg -LogFile $logFile   

        if($storageAccount ){
            Return $storageAccount
        }
        else {
            Return $false
        }
    }
    catch {
        $msg = "ERROR: Failed to find the Azure storage account '$storageAccountName'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit

    }
}

# Function to check if a blob container exists
Function Check-BlobContainer{
    param 
    (      
        [parameter(Mandatory=$true)] [String]$blobContainerName,
        [parameter(Mandatory=$true)] [PSObject]$storageAccount
    )   

    try {
        $result = Get-AzureStorageContainer -Name $blobContainerName -Context $storageAccount.Context -ErrorAction SilentlyContinue

        if($result){
            $msg = "SUCCESS: Blob container '$($blobContainerName)' found under the Storage account '$($storageAccount.StorageAccountName)'."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg -LogFile $logFile  
            Return $true
        }
        else {
            Return $false
        }
    }
    catch {
        $msg = "ERROR: Failed to get the blob container '$($blobContainerName)' under the Storage account '$($storageAccount.StorageAccountName)'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit

    }
}

# Function to create a Blob Container
Function Create-BlobContainer{
    param 
    (      
        [parameter(Mandatory=$true)] [String]$blobContainerName,
        [parameter(Mandatory=$true)] [PSObject]$storageAccount
    )   

    try {
        $result = New-AzureStorageContainer -Name $blobContainerName -Context $storageAccount.Context -ErrorAction Stop

        $msg = "SUCCESS: Blob container '$($blobContainerName)' created under the Storage account '$($storageAccount.StorageAccountName)'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg -LogFile $logFile   
    }
    catch {
        $msg = "ERROR: Failed to create blob container '$($blobContainerName)' under the Storage account '$($storageAccount.StorageAccountName)'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg -LogFile $logFile   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message -LogFile $logFile
        Exit

    }

}

# Function to delete all mailbox connectors
Function Remove-MW_Connectors {

    param 
    (      
        [parameter(Mandatory=$true)] [guid]$CustomerOrganizationId,
        [parameter(Mandatory=$false)] [String]$ProjectType,
        [parameter(Mandatory=$false)] [String]$ProjectName
    )
   
    $connectorPageSize = 100
  	$connectorOffSet = 0
	$connectors = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving $projectType connectors ..."
    
    do
    {   

        if($projectType -eq "Mailbox") {
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $global:mwTicket -OrganizationId $customerOrganizationId -ProjectType "Mailbox" -PageOffset $connectorOffSet -PageSize $connectorPageSize)
        }
        elseif($projectType -eq "Storage"){
            $connectorsPage = @(Get-MW_MailboxConnector -ticket $global:mwTicket -OrganizationId $customerOrganizationId -ProjectType "Storage" -PageOffset $connectorOffSet -PageSize $connectorPageSize)
        }

        if($connectorsPage) {
            $connectors += @($connectorsPage)
            foreach($connector in $connectorsPage) {
                Write-Progress -Activity ("Retrieving connectors (" + $connectors.Length + ")") -Status $connector.Name
            }

            $connectorOffset += $connectorPageSize
        }

    } while($connectorsPage)

    if($connectors -ne $null -and $connectors.Length -ge 1) {
        Write-Host -ForegroundColor Green -Object ("SUCCESS: "+ $connectors.Length.ToString() + " $projectType connector(s) found.") 
    }
    else {
        Write-Host -ForegroundColor Red -Object  "INFO: No $projectType connectors found." 
        Return
    }


    $deletedMailboxConnectorsCount = 0
    $deletedDocumentConnectorsCount = 0
    if($connectors -ne $null) {
        
        Write-Host -ForegroundColor Yellow -Object "INFO: Deleting $projectType connectors:" 

        for ($i=0; $i -lt $connectors.Length; $i++) {
            $connector = $connectors[$i]

            Try {
                if($projectType -eq "Storage"){
                    if($ProjectName -match "FS-DropBox-" -and $connector.Name -match "FS-DropBox-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:mwTicket -Id $connector.Id -force
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif($ProjectName -match "FS-OD4B-" -and $connector.Name -match "FS-OD4B-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:mwTicket -Id $connector.Id -force
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                    elseif($ProjectName -match "Document-" -and $connector.Name -match "Document-") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:mwTicket -Id $connector.Id -force
                        
                        Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                        $deletedDocumentConnectorsCount += 1
                    }
                }    
                
                
                if($projectType -eq "Mailbox") {
                    if($ProjectName -match "Mailbox-O365 Groups conversations" -and $connector.Name -match "Mailbox-O365 Groups conversations") { 
                        $connectorsPage = Remove-MW_MailboxConnector -ticket $global:mwTicket -Id $connector.Id  -force

                         Write-Host -ForegroundColor Green "SUCCESS: $($connector.Name) $projectType connector deleted." 
                         $deletedMailboxConnectorsCount += 1
                    }
                }                         

            }
            catch {
                $msg = "ERROR: Failed to delete $projectType connector $($connector.Name)."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message -LogFile $logFile   
            } 
        }

        
       if($deletedDocumentConnectorsCount -ge 1) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $deletedDocumentConnectorsCount $projectType connector(s) deleted." 
        }
        if($deletedDocumentConnectorsCount -eq 0) {
            if ($projectName -match "FS-OD4B-") {
                Write-Host -ForegroundColor Red "INFO: No $projectType connector was deleted. They were not created by Migrate-MW_AzureBlobContainerToOD4B.ps1."    
            }
            elseif($projectName -match "FS-DropBox-") {
                Write-Host -ForegroundColor Red "INFO: No $projectType connector was deleted. They were not created by Create-MW_AzureBlobContainerToDropBox.ps1."    
            }    
            elseif($projectName -match "Document-") {
                Write-Host -ForegroundColor Red "INFO: No $projectType connector was deleted. They were not created by Create-MW_Office365Groups.ps1."    
            }      
        }

    }

}

# Function to delete all endpoints under a customer
Function Remove-MSPC_Endpoints {
    param 
    (      
        [parameter(Mandatory=$true)] [guid]$customerOrganizationId,
        [parameter(Mandatory=$false)] [String]$endpointType

    )


    $endpointPageSize = 100
  	$endpointOffSet = 0
	$endpoints = $null

    Write-Host
    Write-Host -Object  "INFO: Retrieving MSPC $endpointType endpoints ..."

    $customerTicket = Get-BT_Ticket -OrganizationId $customerOrganizationId

    do {
        
        $endpointsPage = @(Get-BT_Endpoint -Ticket $customerTicket -IsDeleted False -IsArchived False -PageOffset $endpointOffSet -PageSize $endpointPageSize -type $endpointType)

        if($endpointsPage) {
            
            $endpoints += @($endpointsPage)

            foreach($endpoint in $endpointsPage) {
                Write-Progress -Activity ("Retrieving endpoint (" + $endpoints.Length + ")") -Status $endpoint.Name
            }
            
            $endpointOffset += $endpointPageSize
        }
    } while($endpointsPage)

    

    if($endpoints -ne $null -and $endpoints.Length -ge 1) {
        Write-Host -ForegroundColor Green "SUCCESS: $($endpoints.Length) endpoint(s) found."
    }
    else {
        Write-Host -ForegroundColor Red "INFO: No endpoints found." 
    }

    $deletedEndpointsCount = 0

    if($endpoints -ne $null) {
        Write-Host -ForegroundColor Yellow -Object "INFO: Deleting $endpointType endpoints:" 

        for ($i=0; $i -lt $endpoints.Length; $i++) {
            $endpoint = $endpoints[$i]

            Try {
                if($endpoint.Name -match "SRC-" -or $endpoint.Name -match "DST-") {
                    remove-BT_Endpoint -Ticket $customerTicket -Id $endpoint.Id -force
             
                    Write-Host -ForegroundColor Green "SUCCESS: $($endpoint.Name) endpoint deleted." 
                    $deletedEndpointsCount += 1
                }

            }
            catch {
                $msg = "ERROR: Failed to delete endpoint $($endpoint.Name)."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg -LogFile $logFile
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $_.Exception.Message   
            }            
        }

        if($deletedEndpointsCount -ge 1 ) {
            Write-Host
            Write-Host -ForegroundColor Green "SUCCESS: $deletedEndpointsCount $endpointType endpoint(s) deleted." 
        }
        elseif($deletedEndpointsCount -eq 0) {
            Write-Host -ForegroundColor Red "INFO: No $endpointType endpoint was deleted. They were not created by Create-MW_Office365Groups.ps1" 
        }
    }
}

Function Check-FileShare {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$FileShareName,
        [parameter(Mandatory=$true)] [PSObject]$storageAccount
    )   

    try {
        $result = Get-AzureStorageShare -Name $FileShareName -Context $storageAccount.Context -ErrorAction SilentlyContinue

        if($result){
            $msg = "SUCCESS: File Share '$($FileShareName)' found under the Storage account '$($storageAccount.StorageAccountName)'."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg  
            Return $true
        }
        else {
            Return $false
        }
    }
    catch {
        $msg = "ERROR: Failed to get the File Share '$($FileShareName)' under the Storage account '$($storageAccount.StorageAccountName)'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message
        Exit

    }

}

Function Create-FileShare {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$FileShareName,
        [parameter(Mandatory=$true)] [PSObject]$storageAccount
    )   

    try {
        $result = New-AzureStorageShare -Name $FileShareName -Context $storageAccount.Context -ErrorAction Stop

        $msg = "SUCCESS: File Share '$($FileShareName)' created under the Storage account '$($storageAccount.StorageAccountName)'."
        Write-Host -ForegroundColor Green  $msg
        Log-Write -Message $msg   
    }
    catch {
        $msg = "ERROR: Failed to create File Share '$($FileShareName)' under the Storage account '$($storageAccount.StorageAccountName)'. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message
        Exit

    }

    #Calculate the connection strings
    try {
        $AzureSAKey = Get-AzureRmStorageAccountKey -Name $storageAcct.StorageAccountName -ResourceGroupName $storageAcct.ResourceGroupName -ErrorAction Stop | Select-Object -first 1
        
        $StorageAccountKey = $AzureSAKey.Value

        $FileShareUNCPathOutput = "\\$($storageAcct.StorageAccountName).file.core.windows.net\$($result.Name)"

        $FileShareConnectionStringOutput = "net use z: \\$($storageAcct.StorageAccountName).file.core.windows.net\$($result.Name) /u:AZURE\$($storageAcct.StorageAccountName) $($StorageAccountKey)"
    }
    Catch {
        $msg = "ERROR: Failed to get the connection string. Script aborted."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg   
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message
    }


}

Function Create-SASToken {
    param 
    (      
        [parameter(Mandatory=$true)] [String]$blobContainerName,
        [parameter(Mandatory=$true)] [String]$BlobName,
        [parameter(Mandatory=$true)] [PSObject]$storageAccount
    )   

    $StartTime = Get-Date
    $EndTime = $startTime.AddHours(168.0) #1 week

    # Read access - https://docs.microsoft.com/en-us/powershell/module/azure.storage/new-azurestoragecontainersastoken
    $SasToken = New-AzureStorageContainerSASToken -Name $blobContainerName `
    -Context $storageAccount.Context -Permission rl -StartTime $StartTime -ExpiryTime $EndTime
    $SasToken | clip

    # Construnct the URL & Test
    $url = "$($storageAccount.Context.BlobEndPoint)$($blobContainerName)/$($BlobName)$($SasToken)"
    $url | clip

    Return $url

}