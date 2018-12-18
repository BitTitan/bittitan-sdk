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
    Script to     
    1. Generate Office 365 permissions reports
    2. Generate user batches based on FullAccess permissions
    3. Migrate Migrate distribution groups
    4. Migrate mailbox, folder and group permissions to O365
	
.NOTES
	Author			Pablo Galan Sabugo <pablog@bittitan.com> from the Technical Sales Specialist Team <TSTeam@bittitan.com>
	Date		    Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
#>

#######################################################################################################################
#                                               FUNCTIONS
#######################################################################################################################

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
Function Log-Write
{
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}

# Function to get a CSV file name or to create a new CSV file
Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $global:inputFile = $OpenFileDialog.filename

    if($OpenFileDialog.filename -eq "") {
		    # create new import file
	        $inputFileName = "O365Users-import-$((Get-Date).ToString("yyyyMMddHHmmss")).csv"
            $global:inputFile = "$initialDirectory\$inputFileName"

		    #$csv = "primarySmtpAddress`r`n"
		    $file = New-Item -Path $initialDirectory -name $inputFileName -ItemType file -force #-value $csv

            $msg = "SUCCESS: Empty CSV file '$global:inputFile' created."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
            $msg = "WARNING: Populate the CSV file with the email addresses you want to process and save it as`r`n         File Type: 'CSV (Comma Delimited) (.csv)'`r`n         File Name: '$global:inputFile'."
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
			        $importConfirm = $false
		        }
		    }
		    while(-not $importConfirm)
            
            $msg = "SUCCESS: CSV file '$global:inputFile' saved."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
    }
    else{
        $msg = "INFO: CSV file '$($OpenFileDialog.filename)' selected."
        Write-Host  $msg
        Log-Write -Message $msg
    }
}

# Function to query destination email addresses
Function query-EmailAddressMapping {
    do {
        $confirm = (Read-Host -prompt "Are you migrating to the same email addresses?  [Y]es or [N]o")
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($confirm.ToLower() -eq "y") {
        $script:sameEmailAddresses = $true
        $script:sameUserName = $true
        $script:differentDomain = $false
        
        $msg = "WARNING: Since you are migrating the domain to the destination Office 365 tenant,`r`n         either the source or destination primary email addresses must be in onmicrosoft.com format."
        Write-Host -ForegroundColor Yellow $msg      
        
        $script:destinationDomain = (Read-Host -prompt "Please enter the current destination domain")
        $msg = "INFO: Current destination domain is '$script:destinationDomain'."
        Write-Host $msg
        Log-Write -Message $msg
          
    }
    elseif($confirm.ToLower() -eq "n") {
        
        $script:sameEmailAddresses = $false

        do {
            $confirm = (Read-Host -prompt "Are you migrating to a different domain?  [Y]es or [N]o")
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        # If destination Domain is different
        if($confirm.ToLower() -eq "y") {
            $script:differentDomain = $true
            if($createDistributionGroups) {
                do {
                    $script:destinationDomain = (Read-Host -prompt "Please enter the destination domain")
                }while ($script:destinationDomain -eq "")
                $msg = "INFO: Destination domain is '$script:destinationDomain'."
                Write-Host $msg
                Log-Write -Message $msg
            }
            else {
                do{
                    $confirm = (Read-Host -prompt "Are the destination email addresses keeping the same user prefix?  [Y]es or [N]o")
                } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

                if($confirm.ToLower() -eq "y") {
                    $script:sameUserName = $true

                    do {
                        $script:destinationDomain = (Read-Host -prompt "Please enter the destination domain")
                    }while ($script:destinationDomain -eq "")
                     $msg = "INFO: Destination domain is '$script:destinationDomain'."
                     Write-Host $msg
                     Log-Write -Message $msg
                }
                else {
                    $script:sameUserName = $false

                    $msg = "WARNING: Since you are migrating to a different domain`r`n         but you are not keeping the user prefixes the same,`r`n         you will have to manually provide the current destination email addresses."
                    Write-Host -ForegroundColor Yellow $msg     
                    Log-Write -Message $msg
                }    
            }        
        } 
        # If destination domain is the same but user prefix is different, source and destination email addresses must be in onmicrosoft.com format
        else {
            $script:differentDomain = $false
            $script:sameUserName = $false

            
            $msg = "WARNING: Since you are migrating the domain to the destination Office 365 tenant`r`n         but you are not keeping the user prefixes the same,`r`n         you will have to manually provide the current destination email addresses."
            Write-Host -ForegroundColor Yellow $msg     
            Log-Write -Message $msg
        }   
    }
}

#######################################################################################################################
#                                    CONNECTION TO SOURCE AND/OR DESTINATION O365
#######################################################################################################################
# Function to create source EXO PowerShell session
Function Connect-SourceExchangeOnline {
    #Prompt for source Office 365 global admin Credentials
    $msg = "INFO: Connecting to the source Office 365 tenant."
    Write-Host $msg
    Log-Write -Message $msg

    try {
        $loginAttempts = 0
        do {
            $loginAttempts++
            # Connect to Source Exchange Online
            $SourceO365Creds = Get-Credential -Message "Enter Your Source Office 365 Admin Credentials."
            if (!($SourceO365Creds)) {
                $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
                Exit
            }
            $SourceO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $SourceO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue
            $result =Import-PSSession -Session $SourceO365Session -AllowClobber -ErrorAction Stop -WarningAction silentlyContinue -DisableNameChecking -Prefix SRC 
            $msg = "SUCCESS: Connection to source Office 365 Remote PowerShell."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
        }
        until (($loginAttempts -ge 3) -or ($($SourceO365Session.State) -eq "Opened"))

        # Only 3 attempts allowed
        if($loginAttempts -ge 3) {
            $msg = "ERROR: Failed to connect to the Source Office 365. Review your source Office 365 admin credentials and try again."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg
            Start-Sleep -Seconds 5
            Exit
        }
    }
    catch {
        $msg = "ERROR: Failed to connect to source Office 365."
        Write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message
        Get-PSSession | Remove-PSSession
        Exit
    }
    return $SourceO365Session

}

# Function to create destination EXO PowerShell session
Function Connect-DestinationExchangeOnline {
    #Prompt for destination Office 365 global admin Credentials
    $msg = "INFO: Connecting to the destination Office 365 tenant."
    Write-Host $msg
    Log-Write -Message $msg

    try {
        $loginAttempts = 0
        do {
            $loginAttempts++
            # Connect to destination Exchange Online
            $script:destinationO365Creds = Get-Credential -Message "Enter Your Destination Office 365 Admin Credentials."
            if (!($destinationO365Creds)) {
                $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg
                Exit
            }
            $destinationO365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $script:destinationO365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue
            $result =Import-PSSession -Session $destinationO365Session -AllowClobber -ErrorAction Stop -WarningAction silentlyContinue -DisableNameChecking #-Prefix DST 
            $msg = "SUCCESS: Connection to destination Office 365 Remote PowerShell."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg
        }
        until (($loginAttempts -ge 3) -or ($($destinationO365Session.State) -eq "Opened"))

        # Only 3 attempts allowed
        if($loginAttempts -ge 3) {
            $msg = "ERROR: Failed to connect to the destination Office 365. Review your destination Office 365 admin credentials and try again."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg
            Start-Sleep -Seconds 5
            Exit
        }
    }
    catch {
        $msg = "ERROR: Failed to connect to destination Office 365."
        Write-Host -ForegroundColor Red $msg
        Log-Write -Message $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message
        Get-PSSession | Remove-PSSession
        Exit
    }
    return $destinationO365Session
}

#######################################################################################################################
#                                    CHECK IDENTITIES IN DESTINATION O365
#######################################################################################################################

#$recipientList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”,“DiscoveryMailbox”,
#                   “MailContact”,“MailUser”,“GuestMailUser”,
#                   “MailUniversalDistributionGroup”,“MailUniversalSecurityGroup”,“DynamicDistributionGroup”,“RoomList”,
#                   “PublicFolder”)


# Function to check if mailbox exists in destination Office 365
Function check-O365Mailbox {
    param 
    (
        [parameter(Mandatory=$true)] [string]$mailbox
    )

    $recipient = Get-Recipient -identity $mailbox -ErrorAction SilentlyContinue
    
    $mailboxList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”)

    If ($recipient.RecipientType -in $mailboxList -and $recipient.RecipientypeDetails -ne "DiscoveryMailbox") {
        return $true
    }
    else{
        return $false
    }
}

# Function to check if distribution group exists in destination Office 365
Function check-O365Group {
    param 
    (
        [parameter(Mandatory=$true)] [string]$group
    )

    $recipient = Get-Recipient -identity $group -ErrorAction SilentlyContinue
    
    # DynamicDistributionGroup are not supported
    $groupList = @(“MailUniversalDistributionGroup”,“MailUniversalSecurityGroup”,“RoomList”)

    If ($recipient.RecipientType -in $groupList) {
        return $true
    }
    else{
        return $false
    }
}

Function Check-DelegateSource ([string]$DelegateID) {
    
    $CheckDelegate = Get-SRCRecipient $DelegateID -ErrorAction SilentlyContinue

    If ($CheckDelegate -eq $null) {
        $CheckDelegate = Get-SRCGroup $DelegateID -ErrorAction SilentlyContinue 
    }

    If ($CheckDelegate -ne $null) {
        
        If (($ExpandMailSecurityGroups -eq $false -and $CheckDelegate.RecipientType -like "MailUniversalSecurityGroup" ) -or $CheckDelegate.RecipientType -like "*Mailbox") {
            $DelegateName = $CheckDelegate.Name
            $DelegateEmail = $CheckDelegate.PrimarySmtpAddress
            $DelegateAlias = $CheckDelegate.Alias

            Return $DelegateEmail 
        }
        
        # If MailSecurityGroups must be expanded
        If ($ExpandMailSecurityGroups -eq $true -and $CheckDelegate.RecipientType -like "MailUniversalSecurityGroup") {

            $msg = "ALERT: Expand distribution group '$($CheckDelegate.Name)' membership."
            Write-host -ForegroundColor yellow $msg 
            Log-Write -Message $msg
            
            $expandedMembership = @()

            $Members = Get-SRCDistributionGroupMember $CheckDelegate.Name -ResultSize Unlimited
            
            If ($Members){
                ForEach ($Member in $Members) {
                    $CheckMember = Get-SRCRecipient $Member.DistinguishedName -ErrorAction SilentlyContinue
                    If ($CheckMember -ne $null) {
                        $DelegateName = $DelegateID + ":" + $CheckMember.Name
                        $DelegateEmail = $CheckMember.PrimarySmtpAddress
                        $DelegateAlias = $CheckMember.Alias
                        $expandedMembership += $DelegateEmail

                        Return $expandedMembership
                    } 
                } 
            } 
        }
        
        # If Distribution Groups must be expanded
        If ($ExpandBuiltInGroups -eq $true -and $CheckDelegate.RecipientType -eq "Group") {
            
            $msg = "ALERT: Expand built-in group '$($CheckDelegate.Name)' membership."
            Write-host -ForegroundColor yellow $msg 
            Log-Write -Message $msg
            
            $expandedMembership = @()

            $Members = (Get-SRCGroup $DelegateID).Members

            ForEach ($Member in $Members) {
                $CheckMember = Get-SRCRecipient $Member -ErrorAction SilentlyContinue
                
                If ($CheckMember -ne $null) {
                    $DelegateName = $DelegateID + ":" + $CheckMember.Name
                    $DelegateEmail = $CheckMember.PrimarySmtpAddress
                    $DelegateAlias = $CheckMember.Alias
                    $expandedMembership += $DelegateEmail

                    Return $expandedMembership
                } 
            } 
        } 
    }      
    else {
        Return $null
    }

}

#######################################################################################################################
#                                    USER BATCH CREATION BASED ON FULLACCESS PERMISSIONS
#######################################################################################################################

# Function to create batches
Function Create-UserBatches(){
    param(
        [Parameter(Mandatory=$true)]  [array]$InputPermissions
    )
		
    $data = $InputPermissions
    $hashData = $data | Group primarySmtpAddress -AsHashTable -AsString
	$hashDataByDelegate = $data | Group delegateAddress -AsHashTable -AsString
	$usersWithNoDependents = New-Object System.Collections.ArrayList
    $batch = @{}
    $batchCount = 0
    $hashDataSize = $hashData.Count

    $yyyyMMdd = Get-Date -Format 'yyyyMMdd'

	try{
        #Build ArrayList for users with no dependents
        If($hashDataByDelegate["None"].count -gt 0){
		    $hashDataByDelegate["None"] | %{$_.primarySmtpAddress} | %{[void]$usersWithNoDependents.Add($_)}
	    }	    

        #Identify users with no permissions on them, nor them have perms on another
        If($usersWithNoDependents.count -gt 0){
		    $($usersWithNoDependents) | %{
			    if($hashDataByDelegate.ContainsKey($_)){
				    $usersWithNoDependents.Remove($_)
			    }	
		    }
            
            #Remove users with no dependents from hash Data 
            $usersWithNoDependents | %{$hashData.Remove($_)}

		    #Clean out hashData of users in hash data with no delegates, otherwise they'll get batched
		    foreach($key in $($hashData.keys)){
                    if(($hashData[$key] | select -expandproperty delegateAddress ) -eq "None"){
				    $hashData.Remove($key)
			    }
		    }
	    }
        #Execute batch functions
        If(($hashData.count -ne 0) -or ($usersWithNoDependents.count -ne 0)){
            
            while($hashData.count -ne 0) {
                Find-Associations $hashData
            } 

            Write-Host 
            $msg = "INFO: Generating user batches based on FullAcess permissions"
            Write-Host $msg
            Log-Write -Message $msg

            Create-UserBatchFile $batch $usersWithNoDependents   
        }         
    }
    catch {
        $msg = "ERROR: $_"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
    }
}

# Function to identify permission associations    
Function Find-Associations($hashData){
    try{
        #"Hash Data Size: $($hashData.count)" 
        $nextInHash = $hashData.Keys | select -first 1
        $batch.Add($nextInHash,$hashData[$nextInHash])
	
	    Do{
		    $checkForMatches = $false
		    foreach($key in $($hashData.keys)){
	            $Script:comparisonCounter++ 
			
			    Write-Progress -Activity "Analyzing associated delegates" -status "Items remaining: $($hashData.Count)" `
    		    -percentComplete (($hashDataSize-$hashData.Count) / $hashDataSize*100)
			
	            #Checks
			    $usersHashData = $($hashData[$key]) | %{$_.primarySmtpAddress}
                $usersBatch = $($batch[$nextInHash]) | %{$_.primarySmtpAddress}
                $delegatesHashData = $($hashData[$key]) | %{$_.delegateAddress} 
			    $delegatesBatch = $($batch[$nextInHash]) | %{$_.delegateAddress}

			    $ifMatchesHashUserToBatchUser = [bool]($usersHashData | ?{$usersBatch -contains $_})
			    $ifMatchesHashDelegToBatchDeleg = [bool]($delegatesHashData | ?{$delegatesBatch -contains $_})
			    $ifMatchesHashUserToBatchDelegate = [bool]($usersHashData | ?{$delegatesBatch -contains $_})
			    $ifMatchesHashDelegToBatchUser = [bool]($delegatesHashData | ?{$usersBatch -contains $_})
			
			    If($ifMatchesHashDelegToBatchDeleg -OR $ifMatchesHashDelegToBatchUser -OR $ifMatchesHashUserToBatchUser -OR $ifMatchesHashUserToBatchDelegate){
	                if(($key -ne $nextInHash)){ 
					    $batch[$nextInHash] += $hashData[$key]
					    $checkForMatches = $true
	                }
	                $hashData.Remove($key)
	            }
	        }
	    } Until ($checkForMatches -eq $false)
        
        return $hashData 
	}
	catch{
        $msg = "ERROR: $_"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
    }
}

# Function to create batch file
Function Create-UserBatchFile($batchResults,$usersWithNoDepsResults){
	try {
        
        $userBatchesArray = @()
        
	    foreach($key in $batchResults.keys){
            $batchCount++
            $batchName = "BATCH-$batchCount"

		    $output = New-Object System.Collections.ArrayList
		    $($batch[$key]) | %{$output.add($_.primarySmtpAddress)}
		    $($batch[$key]) | %{$output.add($_.delegateAddress)}
                        
            $output | select -Unique | % { $userBatchLineItem = New-Object PSObject;
                                           $userBatchLineItem | Add-Member -MemberType NoteProperty -Name batchName -Value $batchName;
                                           $userBatchLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $_;
                                           $userBatchesArray += $userBatchLineItem }                                                     
            
        }
	    If($usersWithNoDepsResults.count -gt 0){
		     $batchCount++
		     foreach($primarySmtpAddress in $usersWithNoDepsResults){
		 	    $batchName = "BATCH-NoFullAccess"  
              
                $userBatchLineItem = New-Object PSObject
                $userBatchLineItem | Add-Member -MemberType NoteProperty -Name batchName -Value $batchName
                $userBatchLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                $userBatchesArray += $userBatchLineItem 
	        }
	    }
         $msg = "SUCCESS: User batches created: $batchCount" 
         Write-host -ForegroundColor Green $msg 
         Log-Write -Message $msg

         $msg = "         INFO: Number of comparisons: $Script:comparisonCounter" 
         Write-host $msg 
         Log-Write -Message $msg
    }
    catch{
        $msg = "ERROR: $_"
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
    }

    $userBatchesArray | Export-Csv -Path $workingDir\O365UserBatches.csv -NoTypeInformation -force

    $msg = "SUCCESS: CSV file '$workingDir\O365UserBatches.csv' processed, exported and open."
    write-Host -ForegroundColor Green $msg
    Log-Write -Message $msg

    #Open the CSV file for editing
    Start-Process -FilePath $workingDir\O365UserBatches.csv

    $msg = "ACTION: If you have reviewed, edited and saved the CSV file then press any key to continue." 
    Write-Host $msg
    Log-Write -Message $msg
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    
} 

#######################################################################################################################
#                                    EXPORT AND IMPORT O365 DLs & MAIL ENABLED SECURITY GROUPS
#######################################################################################################################
Function Export-O365Groups {
    Try {

        $msg = "INFO: Retrieving groups from Office 365."
        Write-host $msg 
        Log-Write -Message $msg

        $groups = Get-SRCDistributionGroup 

        $totalGroupCount=$groups.count
        $currentGroup=0
        $script:groupsCount=0

        $msg = "SUCCESS: $totalGroupCount groups have been retrieved from Office 365."
        Write-host -ForegroundColor Green $msg 
        Log-Write -Message $msg

        if ($totalGroupCount -eq 0) {
            Exit
        }

        if($createDistributionGroups) {
            foreach ($group in $groups) {
                $currentGroup += 1
                $msg = "INFO: Processing Distribution List $currentGroup/$totalGroupCount : '$($group.DisplayName)' $($group.PrimarySmtpAddress)."
                Write-host $msg 
                Log-Write -Message $msg

	            CreateDL($group)
            }
        }
        
        if($script:groupsCount -ge 2) {
            $msg = "SUCCESS: $groupsCount groups out of $totalGroupCount have been created in destination Office 365."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }elseif ($script:groupsCount -eq 1) {
            $msg = "SUCCESS: 1 FullAccess permission out of $totalGroupCount has been created in destination Office 365."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }

    }
    Catch {
        $msg = "ERROR: Failed to get all O365 groups."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Exit
    }

}

Function CreateDL($group) {

	#$domainName = (Get-MsolDomain | Where-Object { $_.IsDefault -eq $true}).Name

	#Retrieve values of source group 
    $sourceGroupName = $group.DisplayName
    $sourceMembers = Get-SRCDistributionGroupMember $group.Name
    $sourceManagedBy = $group.ManagedBy
		      
	$sourceGroupName = [string]$group.Name
	$sourceGroupDisplayName = [string]$group.DisplayName
	$sourceGroupAlias = [string]$group.Alias	
	$sourceGroupEmailAddresses = $group.EmailAddresses

    
	#rebuilding members, managedby and moderatedby. Check if mailbox exist. If not remove them from list.	
    $destinationPrimarySmtpAddress = [string]$group.PrimarySmtpAddress.Split('@')[0] + "@"+ $script:destinationDomain

	$destinationModeratedBy = New-Object System.Collections.ArrayList
    $destinationMembers = New-Object System.Collections.ArrayList
    $destinationManagedBy = New-Object System.Collections.ArrayList

	foreach ($member in $sourceMembers) {
		if ($member.RecipientType -eq "MailUniversalDistributionGroup" -or $member.RecipientType -eq "MailUniversalSecurityGroup") {
			CreateDL($member)
			$supressOutput = $destinationMembers.Add($member.Name)
		}
		else {
			$result = check-O365Mailbox -mailbox $member.Name       
			if ($result) {
				$supressOutput = $destinationMembers.Add($member.Name)
			}
		}
	}

    foreach ($member in $sourceManagedBy) {
		$result = check-O365Mailbox -mailbox $member
		if ($result) {
			$supressOutput = $destinationManagedBy.Add($member)
		}
    }

    if ($destinationManagedBy.Count -eq 0) {
            $msg = "         WARNING: None of the owners of group $sourceGroupName was found on the destination. Adding administrator as an owner."
            Write-host -ForegroundColor Yellow $msg 
            Log-Write -Message $msg
        $destinationManagedBy += $O365DestinationUsername
    }

	foreach ($member in $sourceModeratedBy) {
		$result = check-O365Mailbox -mailbox $member 
		if ($result) {
			$supressOutput = $destinationModeratedBy.Add($member)
		}
	}

	try {
		#Find Group first. If group exist just continue to next loop
		$result = Get-DistributionGroup -Identity $sourceGroupAlias -ErrorAction SilentlyContinue
		if ($result) {
            $msg = "         WARNING: Group '$sourceGroupName' already exists at destination. Skipping group creation."
            Write-host -ForegroundColor Yellow $msg 
            Log-Write -Message $msg
			Continue
		}
        #This example creates a mail-enabled security group
		if ($group.GroupType -like "*SecurityEnabled*") {
    		$result = New-DistributionGroup `
            -Name $sourceGroupName `
            -Alias $sourceGroupAlias `
            -DisplayName $sourceGroupDisplayName `
            -ManagedBy $destinationManagedBy `
            -Members $destinationMembers `
            -PrimarySmtpAddress $destinationPrimarySmtpAddress -Type Security       
        }
        else {
    		$result = New-DistributionGroup `
            -Name $sourceGroupName `
            -Alias $sourceGroupAlias `
            -DisplayName $sourceGroupDisplayName `
            -ManagedBy $destinationManagedBy `
            -Members $destinationMembers `
            -PrimarySmtpAddress $destinationPrimarySmtpAddress  
            
            #RoomList switch specifies that all members of this distribution group are room mailboxes.
            if($group.RecipientTypeDetails -eq "RoomList") {
                $result = Set-DistributionGroup `
			    -Identity "$sourceGroupName" `
                -RoomList 
            }   
        }

        #to avoid this exception 
        #The "AcceptMessagesOnlyFromSendersOrMembers" and "AcceptMessagesOnlyFrom" parameters can't be specified at the same time.
        $result = Set-DistributionGroup `
			-Identity "$sourceGroupName" `
			-AcceptMessagesOnlyFromSendersOrMembers $group.AcceptMessagesOnlyFromSendersOrMembers `
			-RejectMessagesFromSendersOrMembers $group.RejectMessagesFromSendersOrMembers 

        $result = Set-DistributionGroup `
			-Identity "$sourceGroupName" `
            -AcceptMessagesOnlyFrom $group.AcceptMessagesOnlyFrom `
            -RejectMessagesFrom $group.RejectMessagesFrom 
            		
        #https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/set-distributiongroup?view=exchange-ps 
		$result = Set-DistributionGroup `
				-Identity "$sourceGroupName" `
				-AcceptMessagesOnlyFromDLMembers $group.AcceptMessagesOnlyFromDLMembers `
                -BypassModerationFromSendersOrMembers $group.BypassModerationFromSendersOrMembers `
				-BypassNestedModerationEnabled $group.BypassNestedModerationEnabled `
				-CustomAttribute1 $group.CustomAttribute1 `
				-CustomAttribute2 $group.CustomAttribute2 `
				-CustomAttribute3 $group.CustomAttribute3 `
				-CustomAttribute4 $group.CustomAttribute4 `
				-CustomAttribute5 $group.CustomAttribute5 `
				-CustomAttribute6 $group.CustomAttribute6 `
				-CustomAttribute7 $group.CustomAttribute7 `
				-CustomAttribute8 $group.CustomAttribute8 `
				-CustomAttribute9 $group.CustomAttribute9 `
				-CustomAttribute10 $group.CustomAttribute10 `
				-CustomAttribute11 $group.CustomAttribute11 `
				-CustomAttribute12 $group.CustomAttribute12 `
				-CustomAttribute13 $group.CustomAttribute13 `
				-CustomAttribute14 $group.CustomAttribute14 `
				-CustomAttribute15 $group.CustomAttribute15 `
				-ExtensionCustomAttribute1 $group.ExtensionCustomAttribute1 `
				-ExtensionCustomAttribute2 $group.ExtensionCustomAttribute2 `
				-ExtensionCustomAttribute3 $group.ExtensionCustomAttribute3 `
				-ExtensionCustomAttribute4 $group.ExtensionCustomAttribute4 `
				-ExtensionCustomAttribute5 $group.ExtensionCustomAttribute5 `
				-GrantSendOnBehalfTo $group.GrantSendOnBehalfTo `
				-HiddenFromAddressListsEnabled $True `
				-MailTip $group.MailTip `
				-MailTipTranslations $group.MailTipTranslations `
				-MemberDepartRestriction $group.MemberDepartRestriction `
				-MemberJoinRestriction $group.MemberJoinRestriction `
				-ModeratedBy $destinationModeratedBy `
				-ModerationEnabled $group.ModerationEnabled `
                -RejectMessagesFromDLMembers $group.RejectMessagesFromDLMembers `
				-ReportToManagerEnabled $group.ReportToManagerEnabled `
				-ReportToOriginatorEnabled $group.ReportToOriginatorEnabled `
				-RequireSenderAuthenticationEnabled $group.RequireSenderAuthenticationEnabled `
				-SendModerationNotifications $group.SendModerationNotifications `
				-SendOofMessageToOriginatorEnabled $group.SendOofMessageToOriginatorEnabled `
				-BypassSecurityGroupManagerCheck 

        $msg = "         SUCCESS: Distribution group '$sourceGroupDisplayName' ($($group.RecipientType), $($group.RecipientTypeDetails)) created in the destination Office 365."
        Write-host -ForegroundColor Green $msg 
        Log-Write -Message $msg

        $script:groupsCount += 1

		#Translate the email address of the source group to the destination and add them.
		foreach ($address in $sourceGroupEmailAddresses) {			
            $destinationEmailAddress = [string]$address.Split('@')[0] + "@" + $script:destinationDomain
			
            $msg = "         INFO: Adding email address '$destinationEmailAddress' to distribution group."
            Write-host $msg 
            Log-Write -Message $msg

			$result = Set-DistributionGroup -Identity $sourceGroupName -EmailAddresses @{Add=$destinationEmailAddress}
		}
	}
	catch {
        $msg = "         ERROR: Failed to create distribution group '$sourceGroupDisplayName' in destination Office 365."
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $_.Exception.Message

		Continue
	}
}

#######################################################################################################################
#                                               EXPORT O365 PERMISSIONS
#######################################################################################################################
Function Export-O365Permissions {
    param 
    (
        [parameter(Mandatory=$true)] [boolean]$skipNonExistingUser,
        [parameter(Mandatory=$true)] [boolean]$processSendAs,
        [parameter(Mandatory=$true)] [boolean]$processFullAccess,
        [parameter(Mandatory=$true)] [boolean]$processFolders,
        [parameter(Mandatory=$true)] [boolean]$processSendOnBehalfTo,
        [parameter(Mandatory=$true)] [boolean]$processGroups,
        [parameter(Mandatory=$true)] [boolean]$userBatches,
        [parameter(Mandatory=$true)] [boolean]$readCSVfile
    )

    $ExpandBuiltInGroups = $false
    $ExpandMailSecurityGroups = $false

    Try {
        Write-host 
        $msg = "INFO: Retrieving mailboxes from Office 365."
        Write-host $msg 
        Log-Write -Message $msg

       if($readCSVfile) { 

            #Read CSV file
            try {
                $mailboxesInCSV = get-content $global:inputFile
            }
            catch {
                $msg = "ERROR: Failed to import the CSV file '$inputFile'."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $msg
                Log-Write -Message $_.Exception.Message
            }            
            $mailboxes = @($mailboxesInCSV | % { Get-SRCMailbox $_ -ErrorAction silentlyContinue  } | Where-Object {$_.RecipientTypeDetails -ne "DiscoveryMailbox"} | Select DisplayName,UserPrincipalName,PrimarySmtpAddress,RecipientType,RecipientTypeDetails,Identity)
       
        }
        else {
            $mailboxes = @(Get-SRCMailbox -ResultSize Unlimited -ErrorAction Stop| Where-Object {$_.RecipientTypeDetails -ne "DiscoveryMailbox"} | Select DisplayName,UserPrincipalName,PrimarySmtpAddress,RecipientType,RecipientTypeDetails,Identity)
        }

        $totalMailboxCount=$mailboxes.count
        $msg = "SUCCESS: $totalMailboxCount mailboxes have been retrieved from Office 365."
        Write-host -ForegroundColor Green $msg 
        Log-Write -Message $msg

        if($createUserBatches -eq $false -and $processGroups -eq $true) { 
            $distributionGroups = @(Get-SRCDistributionGroup -ResultSize Unlimited -ErrorAction Stop| Select DisplayName,PrimarySmtpAddress,RecipientType,RecipientTypeDetails,Identity)
            $totalGroupCount=$distributionGroups.count
            $msg = "SUCCESS: $totalGroupCount groups have been retrieved from Office 365."
            Write-host -ForegroundColor Green $msg 
            Log-Write -Message $msg

            $mailboxes += $distributionGroups
        }
        
    }
    Catch {
        $msg = "ERROR: Failed to get all O365 mailboxes."
        Log-Write -Message $msg
        Exit
    }

    if($totalMailboxCount -eq 0) {
        Exit
    }

    $msg = "INFO: Exporting all permissions."
    Write-host $msg 
    Log-Write -Message $msg

    #Declare Output Arrays
    $script:fullAccessPermissionsArray = @()
    $script:sendAsPermissionsArray = @()
    $script:sendOnBehalfToPermissionsArray = @()
    $script:folderPermissionsArray = @()
    $userBatchesArray = @()

    $currentMailbox=0
    $currentGroup=0

    $mailboxList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”)
    $groupList = @(“MailUniversalDistributionGroup”,“MailUniversalSecurityGroup”,“RoomList”)
    
    #Process each source mailbox to export permissions
    foreach ($mailbox in $mailboxes) {

        ######################################################## 
        # Current mailbox 
        ######################################################## 
        $displayName = $mailbox.DisplayName
        $identity = $mailbox.Identity
        $primarySmtpAddress = $mailbox.PrimarySmtpAddress
        $recipientType = $mailbox.RecipientType
        $recipientTypeDetails = $mailbox.RecipientTypeDetails

        if($recipientType -in $mailboxList ) {

            $name = $mailbox.UserPrincipalName

            $currentMailbox += 1
            $msg = "INFO: Processing mailbox $currentMailbox/$totalMailboxCount : '$displayName' $primarySmtpAddress ($recipientType, $recipientTypeDetails)."
            Write-host $msg 
            Log-Write -Message $msg

            # Skip the user process only if the 3 conditions below are met:
            # 1. Importing permissions into destination Office 365 (not only exporting permissions from O365)
            # 2. The destination Office 365 email addresses do not have different userName and different domain
            # 3. The user does not exist in destination Office 365
            if($skipNonExistingUser) {
                if ($script:sameEmailAddresses -and $script:destinationDomain -ne "") {
                    #From O365 to O365 migrating the domain, we need to use onmicrosoft.com either at source or destination, so we need the current destination domain:
                    $destinationPrimarySmtpAddress = $primarySmtpAddress.Split('@')[0] + "@" + $script:destinationDomain
                }
                elseif (!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "") {
                    $destinationPrimarySmtpAddress = $primarySmtpAddress.Split('@')[0] + "@" + $script:destinationDomain
                } 
                elseif (!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "") {
                    #Since we don't have the destination primary SMTP address druing the mailbox export, will try with the displayName
                    $destinationPrimarySmtpAddress = $displayName
                }

                $result=check-O365Mailbox -mailbox $destinationPrimarySmtpAddress    

                if($result){
                    $msg = "      SUCCESS: User '$displayName' $destinationPrimarySmtpAddress found in destination Office 365 tenant."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg
                }
                else {
                    $msg = "      INFO: User '$displayName' $destinationPrimarySmtpAddress not found in destination Office 365 tenant."
                    Write-Host -ForegroundColor Red $msg
                    Log-Write -Message $msg
                    $msg = "      Skipping user processing."
                    Write-Host -ForegroundColor Red $msg
                    Log-Write -Message $msg
                    Continue
                }
            }
        }

        if($recipientType -in $groupList -and ($processSendAs -eq $true -or $processSendOnBehalfTo -eq $true)) {

            $name = $mailbox.PrimarySmtpAddress

            $currentGroup += 1
            $msg = "INFO: Processing group $currentGroup/$totalGroupCount : '$displayName' $primarySmtpAddress ($recipientType,$recipientTypeDetails)."
            Write-host $msg 
            Log-Write -Message $msg

            # Skip the user process only if the 3 conditions below are met:
            # 1. Importing permissions into destination Office 365 (not only exporting permissions from O365)
            # 2. The destination Office 365 email addresses do not have different userName and different domain
            # 3. The user does not exist in destination Office 365
            if($skipNonExistingUser) {
                if ($script:sameEmailAddresses -and $script:destinationDomain -ne "") {
                    #From O365 to O365 migrating the domain, we need to use onmicrosoft.com either at source or destination, so we need the current destination domain:
                    $destinationPrimarySmtpAddress = $primarySmtpAddress.Split('@')[0] + "@" + $script:destinationDomain
                }
                elseif (!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "") {
                    $destinationPrimarySmtpAddress = $primarySmtpAddress.Split('@')[0] + "@" + $script:destinationDomain
                } 
                elseif (!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "") {
                    #Since we don't have the destination primary SMTP address druing the mailbox export, will try with the displayName
                    $destinationPrimarySmtpAddress = $displayName
                }

                $result=check-O365Group -group $destinationPrimarySmtpAddress    

                if($result){
                    $msg = "      SUCCESS: Group '$displayName' $destinationPrimarySmtpAddress found in destination Office 365 tenant."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg
                }
                else {
                    $msg = "      INFO: Group '$displayName' $destinationPrimarySmtpAddress not found in destination Office 365 tenant."
                    Write-Host -ForegroundColor Red $msg
                    Log-Write -Message $msg
                    $msg = "      Skipping group processing."
                    Write-Host -ForegroundColor Red $msg
                    Log-Write -Message $msg
                    Continue
                }
            }
        }



        ######################################################### 
	    # Retrieve SendAs permissions	
        ########################################################  
	    if($processSendAs) {
            if($recipientType -in $mailboxList ) {
                $sendAsPermissions = @(Get-SRCMailbox -Identity $name | Get-SRCRecipientPermission -ResultSize Unlimited | where {$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -ne “NULL SID” -and $_.IsInherited -eq $false} | select Identity,Trustee,AccessRights, AccessControlType, IsInherited, InheritanceType)
            }
            elseif($recipientType -in $groupList ) {
                $sendAsPermissions = @(Get-SRCDistributionGroup -Identity $name | Get-SRCRecipientPermission -ResultSize Unlimited | where {$_.Trustee -ne "NT AUTHORITY\SELF" -and $_.Trustee -ne “NULL SID” -and $_.IsInherited -eq $false} | select Identity,Trustee,AccessRights, AccessControlType, IsInherited, InheritanceType)
            }

            $msg = "      INFO: $($sendAsPermissions.count) SendAs permissions have been found for '$displayName'."
            Write-host $msg 
            Log-Write -Message $msg

            foreach($sendAsPermission in $sendAsPermissions) {
                $displayName
                $primarySmtpAddress     
                $recipientType   
                $recipientTypeDetails   
                $trustee = $sendAsPermission.Trustee
                $accessRights = $sendAsPermission.AccessRights
                $accessControlType = $sendAsPermission.AccessControlType        
                $isInherited = $sendAsPermission.IsInherited
                $inheritanceType = $sendAsPermission.InheritanceType
            
                if($script:destinationDomain -ne "") {
                    $userSplit = $primarySmtpAddress -split "@"
                    $userName = $userSplit[0]
                    $userDomain = $userSplit[1]
                    $destinationPrimarySmtpAddress = "$userName@$script:destinationDomain" 

                    $sendAsEmailSplit = $trustee -split "@"
                    $sendAsUserName = $sendAsEmailSplit[0]
                    $sendAsDomain = $sendAsEmailSplit[1]
                    $destinationTrustee = "$sendAsUserName@$script:destinationDomain"                 
                }

                $sendAsLineItem = New-Object PSObject

                #Same domain and same user prefix (either source or destination must be in onmicrosoft.com format)
                if ($script:sameEmailAddresses -and !$script:differentDomain -and $script:sameUserName) {
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                    if($script:destinationDomain -ne "") {
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value $destinationPrimarySmtpAddress
                    }
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails

                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name trustee -Value $trustee  
                    if($script:destinationDomain -ne "") {
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationTrustee -Value $destinationTrustee
                    }

                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name accessRights -Value $accessRights
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name accessControlType -Value $accessControlType
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name isInherited -Value $isInherited
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name inheritanceType -Value $inheritanceType  
                }
                #Different domain but same user prefix 
                elseif (!$script:sameEmailAddresses -and $script:differentDomain -and $script:sameUserName) {                    
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value $destinationPrimarySmtpAddress
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails
                        
                    #if userDomain and sendAsDomain are the same in Office 365, sendAsEmail domain is automatically changed to the new domain
                    if($userDomain -eq $sendAsDomain) {
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name trustee -Value $destinationTrustee
                    }
                    #if userDomain and sendAsDomain are not the same in Office 365, sendAsEmail domain must be entered manually by the user in the CSV file
                    else {
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name trustee -Value $trustee  
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationTrustee -Value "" 
                    }   

                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name accessRights -Value $accessRights
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name accessControlType -Value $accessControlType
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name isInherited -Value $isInherited
                    $sendAsLineItem | Add-Member -MemberType NoteProperty -Name inheritanceType -Value $inheritanceType  	
                }
                #Different user prefix, destination email addresses must be manually entered in the CSV file 
                elseif(!$script:sameEmailAddresses -and !$script:sameUserName) {  
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value ""
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails

                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name trustee -Value $trustee  
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name destinationTrustee -Value "" 

                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name accessRights -Value $accessRights
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name accessControlType -Value $accessControlType
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name isInherited -Value $isInherited
                        $sendAsLineItem | Add-Member -MemberType NoteProperty -Name inheritanceType -Value $inheritanceType  	
                }

                $script:sendAsPermissionsArray += $sendAsLineItem
            }
	    }

        ######################################################## 
	    # Retrieve Full Access permissions	
        ########################################################   
	    if($processFullAccess) {   
            if($recipientType -in $mailboxList ) {
            #Do not include inherited permissions. Only explicit permissions are migrated 
            $fullAccessPermissions = @(Get-SRCMailbox -Identity $name | Get-SRCMailboxPermission -ResultSize Unlimited | Where-Object {$_.user -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false} | select Identity,User,AccessRights)
       
            if($($fullAccessPermissions.count) -eq 0) {
                $fullAccessLineItem = New-Object PSObject
                $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value "None"
                $userBatchesArray += $fullAccessLineItem
            }
	    
            $msg = "      INFO: $($fullAccessPermissions.count) FullAcess permissions have been found for '$displayName'."
            Write-host $msg 
            Log-Write -Message $msg
                
            foreach($fullAccessPermission in $fullAccessPermissions) {    
                $displayName
                $primarySmtpAddress
                $delegateAddress = Check-DelegateSource $fullAccessPermission.User
                $accessRights = $fullAccessPermission.AccessRights
                $recipientType

                if($script:destinationDomain -ne "") {
                    $userSplit = $primarySmtpAddress -split "@"
                    $userName = $userSplit[0]
                    $userDomain = $userSplit[1]
                    $destinationPrimarySmtpAddress = "$userName@$script:destinationDomain" 

                    $delegateSplit = $delegateAddress -split "@"
                    $delegateUserName = $delegateSplit[0]
                    $delegateDomain = $delegateSplit[1]
                    $destinationDelegateAddress= "$delegateUserName@$script:destinationDomain"                 
                }
    
                $fullAccessLineItem = New-Object PSObject

                #Same domain and same user prefix (either source or destination must be in onmicrosoft.com format)
                if ($script:sameEmailAddresses -and !$script:differentDomain -and $script:sameUserName) {
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                    if($script:destinationDomain -ne "") {
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value $destinationPrimarySmtpAddress
                    }
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType	
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails 

                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                    if($script:destinationDomain -ne "") {
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name destinationDelegateAddress -Value $destinationDelegateAddress
                    }
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name accessRights -Value $accessRights    
                    #AutoMapping $true by default
                    $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name autoMapping -Value $true              

                }
                #Different domain but same user prefix 
                elseif (!$script:sameEmailAddresses -and $script:differentDomain -and $script:sameUserName) {
                            
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $destinationPrimarySmtpAddress
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType	
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails 

                        #if userDomain and fullAccessDomain are the same in Office 365, delegate domain is automatically changed to the new domain
                        if($userDomain -eq $delegateDomain) {
                            $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $destinationDelegateAddress
                        }
                        #if userDomain and fullAccessDomain are not the same in Office 365, delegate domain must be entered manually by the user in the CSV file
                        else {
                            $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                            $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name destinationDelegateAddress -Value ""
                        }
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name accessRights -Value $accessRights     
                        #AutoMapping $true by default
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name autoMapping -Value $true             
                }
                #Different user prefix, destination email addresses must be manually entered in the CSV file 
                elseif(!$script:sameEmailAddresses -and !$script:sameUserName) {  

                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value ""
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType	 
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails	 

                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name destinationDelegateAddress -Value ""
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name accessRights -Value $accessRights  
                        #AutoMapping $true by default
                        $fullAccessLineItem | Add-Member -MemberType NoteProperty -Name autoMapping -Value $true   
   
                }

                $script:fullAccessPermissionsArray += $fullAccessLineItem   
                $userBatchesArray += $fullAccessLineItem 
            }
            }
        }

        ######################################################### 
        # Retrieve mailbox folder permissions
        ######################################################### 
	    if($processFolders) {

            if($recipientType -in $mailboxList ) {
     
            if($onlyPermissionsReport -and !$processOnlyCalendars) {
                $mailboxFolders = @(Get-SRCMailbox -Identity $name | Get-SRCMailboxFolderStatistics)
            }
            else {
                $mailboxFolders = @(Get-SRCMailbox -Identity $name | Get-SRCMailboxFolderStatistics -FolderScope Calendar)
	        }
            $folderPermissionsCount = 0
            $calendarPermissionsCount = 0

            # Loop through each user who has Folder Delegate Permissions and find its corresponding user on the destination and set the Folder Permissions
	        foreach($mailboxFolder in $mailboxFolders) {
		        $folderPath = $mailboxFolder.FolderPath.Replace("/","\")
		        if ($folderPath -eq "\Top of the Information Store") {
			        $folderPath = "\"
		        }
		        $folderLocation = $name + ":" + $folderPath
                $folderType = $mailboxFolder.FolderType

		        $folderPermissions = @(Get-SRCMailboxFolderPermission -Identity $folderLocation -ErrorAction SilentlyContinue)
		    
                if ($folderPermissions -ne $null) {
                
			        foreach ($folderPermission in $folderPermissions) {
				        [string]$folderDelegate = $folderPermission.User

                        $delegate = Get-SRCRecipient $folderDelegate -ErrorAction SilentlyContinue | select Identity,PrimarySmtpAddress,RecipientType,RecipientTypeDetails 
					    if ($delegate -ne $null) {

                            $delegateIdentity = $delegate.Identity
						    $delegateAddress = $delegate.PrimarySmtpAddress
						    $delegateAccess = $folderPermission.AccessRights
                            $folderName = $folderPermission.FolderName
                            $folderIdentity = $mailbox.PrimarySmtpAddress + ":\" + $folderName

				            If ($delegateIdentity -ne $identity -and $folderDelegate -ne "Default" -and $folderDelegate -ne "Anonymous" -and $delegateAccess -ne "None") {

                                $folderPermissionsCount += 1
                                if($folderType -eq “Calendar”) {
                                    $calendarPermissionsCount += 1                            
                             
                                }       
                                
                                if($script:destinationDomain -ne "") {
                                    $userSplit = $primarySmtpAddress -split "@"
                                    $userName = $userSplit[0]
                                    $userDomain = $userSplit[1]
                                    $destinationPrimarySmtpAddress = "$userName@$script:destinationDomain" 

                                    $delegateAddressSplit = $delegateAddress -split "@"
                                    $delegateUserName = $delegateAddressSplit[0]
                                    $delegateDomain = $delegateAddressSplit[1]
                                    $destinationDelegateAddress = "$delegateUserName@$script:destinationDomain" 
                                }

                                $folderPermissionLineItem = New-Object PSObject

                                #Same domain and same user prefix (either source or destination must be in onmicrosoft.com format)
                                if ($script:sameEmailAddresses -and !$script:differentDomain -and $script:sameUserName) {

                                    $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                                    $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                                    if($script:destinationDomain -ne "") {
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value $destinationPrimarySmtpAddress
                                    }
                                    $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType
                                    $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails

                                    $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                                    if($script:destinationDomain -ne "") {
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name destinationDelegateAddress -Value $destinationDelegateAddress
                                    }
                                    $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name delegateAccess -Value $delegateAccess
                                    $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name folderName -Value $folderName	 
                                    $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name folderType -Value $folderType   
                                    
                                }
                                #Different domain but same user prefix 
                                elseif (!$script:sameEmailAddresses -and $script:differentDomain -and $script:sameUserName) {

                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $destinationPrimarySmtpAddress
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails
                                    
                                        #if userDomain and folderPermissionDomain are the same in Office 365, folderPermissionEmail domain is automatically changed to the new domain
                                        if($userDomain -eq $delegateDomain) {
                                            $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $destinationDelegateAddress
                                        }
                                        #if userDomain and folderPermissionDomain are not the same in Office 365, folderPermissionEmail domain must be entered manually by the user in the CSV file
                                        else {
                                            $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                                            $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name destinationDelegateAddress -Value ""
                                        }                
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name delegateAccess -Value $delegateAccess
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name folderName -Value $folderName	 
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name folderType -Value $folderType 
                                               }
                                #Different user prefix, destination email addresses must be manually entered in the CSV file 
                                elseif(!$script:sameEmailAddresses -and !$script:sameUserName) {       

                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value ""
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails

                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name delegateAddress -Value $delegateAddress
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name destinationDelegateAddress -Value ""
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name DelegateAccess -Value $delegateAccess
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name folderName -Value $folderName	 
                                        $folderPermissionLineItem | Add-Member -MemberType NoteProperty -Name folderType -Value $folderType     
                                    
                                }

                                $script:folderPermissionsArray += $folderPermissionLineItem
				            }
                        }
			        }
		        }            

	        }
        
            $msg = "      INFO: $folderPermissionsCount folder permissions ($calendarPermissionsCount calendar permissions) have been found for '$displayName'."
            Write-host $msg 
            Log-Write -Message $msg
            }
        }

        ######################################################### 
	    # Retrieve SendOnBehalfTo permissions	
        ########################################################   
        if($processSendOnBehalfTo){
            if($recipientType -in $mailboxList ) {
                $sendOnBehalfToPermissions = @(Get-SRCMailbox -Identity $name | Where-Object {$_.GrantSendOnBehalfTo -ne $null} | select GrantSendOnBehalfTo,PrimarySmtpAddress,RecipientType,RecipientTypeDetails)
            }
            elseif($recipientType -in $groupList ) {
                $sendOnBehalfToPermissions = @(Get-SRCDistributionGroup -Identity $name | Where-Object {$_.GrantSendOnBehalfTo -ne $null} | select GrantSendOnBehalfTo,PrimarySmtpAddress,RecipientType,RecipientTypeDetails)
            }

            $msg = "      INFO: $($sendOnBehalfToPermissions.GrantSendOnBehalfTo.count) SendOnBehalfTo permissions have been found for '$displayName'."
            Write-host $msg 
            Log-Write -Message $msg

            $displayName
            $primarySmtpAddress #= $sendOnBehalfToPermissions.PrimarySmtpAddress
            $recipientType = $sendOnBehalfToPermissions.RecipientType
            $recipientTypeDetails = $sendOnBehalfToPermissions.RecipientTypeDetails

            foreach($sendOnBehalfToPermission in $sendOnBehalfToPermissions.GrantSendOnBehalfTo) {    
                $SendOnBehalfToLineItem = New-Object PSObject

                #Populate destinationPrimaryEmail only when migrating to the same email address
                if($script:destinationDomain -ne "") {
                     $userSplit = $primarySmtpAddress -split "@"
                     $userName = $userSplit[0]
                     $userDomain = $userSplit[1]
                     $destinationPrimarySmtpAddress = "$userName@$script:destinationDomain" 
                }

                #Same domain and same user prefix (either source or destination must be in onmicrosoft.com format)
                if ($script:sameEmailAddresses -and !$script:differentDomain -and $script:sameUserName) {
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                    if($script:destinationDomain -ne "") {
                        $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value $destinationPrimarySmtpAddress
                    }
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType 
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails

                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name GrantSendOnBehalfTo -Value $sendOnBehalfToPermission  
                }
                #Different domain but same user prefix 
                elseif (!$script:sameEmailAddresses -and $script:differentDomain -and $script:sameUserName) {

                     $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                     $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $destinationPrimarySmtpAddress
                     $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType 
                     $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails
     
                     $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name grantSendOnBehalfTo -Value $sendOnBehalfToPermission             
                }
                #Different user prefix, destination email addresses must be manually entered in the CSV file 
                elseif(!$script:sameEmailAddresses -and !$script:sameUserName) {  
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name displayName -Value $displayName
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name primarySmtpAddress -Value $primarySmtpAddress
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name destinationPrimarySmtpAddress -Value ""
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name recipientType -Value $recipientType 
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name recipientTypeDetails -Value $recipientTypeDetails

                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name grantSendOnBehalfTo -Value $sendOnBehalfToPermission
                    $SendOnBehalfToLineItem | Add-Member -MemberType NoteProperty -Name destinationGrantSendOnBehalfTo -Value ""       
                 }

                $script:sendOnBehalfToPermissionsArray += $SendOnBehalfToLineItem
            }

        }
    }
    
    if($createUserBatches -eq $true -and $userBatchesArray -ne $null){
        Create-UserBatches -InputPermissions $userBatchesArray
    }
}

#######################################################################################################################
#                                        IMPORT SENDAS PERMISSIONS INTO O365
#######################################################################################################################
Function Process-SendAsPermissions {
    

    $msg = "INFO: Exporting SendAs permissions to CSV file."
    Write-Host $msg
    Log-Write -Message $msg

    if($script:sendAsPermissionsArray -ne $null) { 
    
    #Export sendAsPermissionsArray to CSV file
    try {
        if($onlyPermissionsReport) {
            $script:sendAsPermissionsArray | Export-Csv -Path $workingDir\O365SendAsPermissionsReport.csv -NoTypeInformation -force
            $msg = "SUCCESS: CSV file '$workingDir\O365SendAsPermissionsReport.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
        else {
            $script:sendAsPermissionsArray | Export-Csv -Path $workingDir\O365SendAsPermissions.csv -NoTypeInformation -force
            $msg = "SUCCESS: CSV file '$workingDir\O365SendAsPermissions.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }  

        if ($script:sameEmailAddresses) {
            $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "" -and $userDomain -eq $sendAsDomain -and $onlyPermissionsReport -eq $false) {
            $msg = "         WARNING: The 'primarySmtpAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         WARNING: The 'sendAsEmail' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""  -and $userDomain -ne $sendAsDomain -and $onlyPermissionsReport -eq $false) {
            $msg = "         WARNING: The 'primarySmtpAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Populate the 'destinationSendAsEmail' column of the opened CSV file with the destination SendAs email."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "" -and $onlyPermissionsReport -eq $false) {
            $msg = "         ACTION:  Populate the 'destinationUser' column of the opened CSV file with the destination user email."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Populate the 'destinationSendAsEmail' column of the opened CSV file with the destination SendAs email."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Once you finish editing the CSV file, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif ($onlyPermissionsReport -eq $false) {
        }
    }
    catch {
        if($onlyPermissionsReport) {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365SendAsPermissionsReport.csv'."
        }
        else {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365SendAsPermissions.csv'."
        } 
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg
        Log-Write -Message $_.Exception.Message
        Exit
    }

    #Open the CSV file for editing
    if($onlyPermissionsReport) {
        Start-Process -FilePath $workingDir\O365SendAsPermissionsReport.csv
    }
    else {
        Start-Process -FilePath $workingDir\O365SendAsPermissions.csv
    }    

    #If the script must generate GSuite permissions report and also migrate them to O365
    if(!$onlyPermissionsReport) {
        $msg = "ACTION:  If you have reviewed, edited and saved the CSV file then press any key to continue." 
        Write-Host $msg
        Log-Write -Message $msg
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

        #Re-import the edited CSV file
        Try{
            $sendAsPermissions = @(Import-CSV "$workingDir\O365SendAsPermissions.csv" | where-Object { $_.PSObject.Properties.Value -ne ""})
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365SendAsPermissions.csv'."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        $totalSendAsPermissionsExport = $sendAsPermissions.Count
        $sendAsPermissionsCount = 0
        $currentSendAsPermission = 0

        if($totalSendAsPermissionsExport -eq 0) {
            Write-Host 
            $msg = "INFO: No SendAs permissions found in exported CSV file."  
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg  
 
        }
        else {
            Write-Host 
            $msg = "INFO: Importing SendAs permissions into Office 365."
            Write-Host $msg
            Log-Write -Message $msg

            Foreach($sendAsPermission in $sendAsPermissions){
                #Current SendAs Permission
                $currentSendAsPermission += 1

                $displayName = $sendAsPermission.DisplayName
                if($sendAsPermission.DestinationPrimarySmtpAddress) {
                    $targetMailbox = $sendAsPermission.DestinationPrimarySmtpAddress 
                }
                else {
                    $targetMailbox = $sendAsPermission.PrimarySmtpAddress 
                }
                if($sendAsPermission.DestinationTrustee) {
                    $sendAsEmail = $sendAsPermission.DestinationTrustee
                }
                else {
                    $sendAsEmail = $sendAsPermission.Trustee
                }
                            
                $recipientType = $sendAsPermission.RecipientType   
                $accessRights = $accessRights = $sendAsPermission.AccessRights
                $accessControlType = $accessControlType = $sendAsPermission.AccessControlType        
                $isInherited = $isInherited = $sendAsPermission.IsInherited
                $inheritanceType = $inheritanceType = $sendAsPermission.InheritanceType

                $mailboxList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”)
                $groupList = @(“MailUniversalDistributionGroup”,“MailUniversalSecurityGroup”,“RoomList”)
                
                if($($sendAsPermission.RecipientType) -in $mailboxList ) {            
                    $msg = "INFO: Processing SendAs permission $currentSendAsPermission/$totalSendAsPermissionsExport : TargetMailbox $targetMailbox SendAsEmail $sendAsEmail."
                    Write-host $msg
                    Log-Write -Message $msg

                    #Verify if target mailbox exists    
                    $recipient = check-O365Mailbox -mailbox $targetMailbox
                }
                elseif($($sendAsPermission.RecipientType) -in $groupList) {
                    $msg = "INFO: Processing SendAs permission $currentSendAsPermission/$totalSendAsPermissionsExport : TargetGroup $targetMailbox SendAsEmail $sendAsEmail."
                    Write-host $msg
                    Log-Write -Message $msg

                    #Verify if target group exists    
                    $recipient = check-O365Group -group $targetMailbox
                }
            
                If ($recipient -eq $true) {

                    #Verify if sendAsEmail exists
                    $recipient = check-O365Mailbox -mailbox $sendAsEmail 

                    If($recipient -eq $true) {
                        try {
                            $result = Get-RecipientPermission $targetMailbox -Trustee $sendAsEmail -AccessRights "SendAs"
                            if (!$result) {
                                $result=Add-RecipientPermission $targetMailbox -Trustee $sendAsEmail -AccessRights "SendAs" -Confirm:$false 

                                $msg = "      SUCCESS: SendAs permission applied."
                                Write-Host -ForegroundColor Green $msg
                                Log-Write -Message $msg
                                $sendAsPermissionsCount += 1
                            }
                            else {
                                $msg = "      WARNING: SendAs permission already exists in Office 365."
                                Write-Host -ForegroundColor Yellow $msg 
                                Log-Write -Message $msg
                            }
                        }
                        catch {
                            $msg = "      ERROR: Failed to apply SendAs permission."
                            Write-Host -ForegroundColor Red  $msg
                            Write-Host -ForegroundColor Red $_.Exception.Message
                            Log-Write -Message $msg
                            Log-Write -Message $_.Exception.Message
                        }
                    }
                    else {
                        $msg = "      ERROR: SendAsEmail '$sendAsEmail' doest not exist in Office 365. SendAs permission skipped."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg
                    }        
                } 
                else{
                    $msg =  "      ERROR: Target mailbox '$targetMailbox' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                }  


            }

            if($sendAsPermissionsCount -ge 2) {
                $msg = "SUCCESS: $sendAsPermissionsCount SendAs permissions out of $totalSendAsPermissionsExport have been applied to Office 365 mailboxes."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
            } elseif ($sendAsPermissionsCount -eq 1) {
                $msg = "SUCCESS: 1 SendAs permission out of $totalSendAsPermissionsExport has been applied to Office 365 mailboxes."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
            }
        }
    }
    }
    else {
        $msg = "INFO: No SendAs permissions found in Office 365." 
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg  
    }
}

#######################################################################################################################
#                                        IMPORT FULLACCESS PERMISSIONS INTO O365
#######################################################################################################################
Function Process-FullAccessPermissions {

    $msg = "INFO: Exporting FullAccess permissions to CSV file."
    Write-Host $msg
    Log-Write -Message $msg

    if($script:fullAccessPermissionsArray -ne $null) { 
    #Export fullAccessPermissionsArray to CSV file
    try {
        if($onlyPermissionsReport) {
            $script:fullAccessPermissionsArray | Export-Csv -Path $workingDir\O365FullAccessPermissionsReport.csv -NoTypeInformation -force
            $msg = "SUCCESS: CSV file '$workingDir\O365FullAccessPermissionsReport.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }
        else {
            $script:fullAccessPermissionsArray | Export-Csv -Path $workingDir\O365FullAccessPermissions.csv -NoTypeInformation -force
            $msg = "SUCCESS: CSV file '$workingDir\O365FullAccessPermissions.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg
        }          
        if ($script:sameEmailAddresses) {
            $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow  $msg
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "" -and $userDomain -eq $fullAccessDomain -and $onlyPermissionsReport -eq $false) {
            $msg = "         WARNING: The 'primarySmtpAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         WARNING: The 'delegateAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""  -and $userDomain -ne $fullAccessDomain -and $onlyPermissionsReport -eq $false) {
            $msg = "         WARNING: The 'primarySmtpAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Populate the 'destinationDelegateAddress' column of the opened CSV file with the destination FullAccess email."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "" -and $onlyPermissionsReport -eq $false) {
            $msg = "         WARNING: Populate the 'destinationUser' column of the opened CSV file with the destination user email."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         WARNING: Populate the 'destinationDelegateAddress' column of the opened CSV file with the destination FullAccess email."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Once you finish editing the CSV file, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif($onlyPermissionsReport -eq $false) {
        }
    }
    catch {
        if($onlyPermissionsReport) {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365FullAccessPermissionsReport.csv'."
        }
        else {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365FullAccessPermissions.csv'."
        } 
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg
        Log-Write -Message $_.Exception.Message
        Exit
    }

    #Open the CSV file for editing
    if($onlyPermissionsReport) {
        Start-Process -FilePath $workingDir\O365FullAccessPermissionsReport.csv
    }
    else {
        Start-Process -FilePath $workingDir\O365FullAccessPermissions.csv
    }    
    
    #If the script must generate GSuite permissions report and also migrate them to O365
    if(!$onlyPermissionsReport) {
        $msg = "ACTION: If you have reviewed, edited and saved the CSV file then press any key to continue."
        Write-Host $msg
        Log-Write -Message $msg
        
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

        #Re-import the edited CSV file
        Try{
            $fullAccessPermissions = @(Import-CSV "$workingDir\O365FullAccessPermissions.csv" | where-Object { $_.PSObject.Properties.Value -ne ""})
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365FullAccessPermissions.csv'. Please save and close the CSV file."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        $confirm = (Read-Host -prompt "Do you want to enable the auto-mapping feature in Microsoft Outlook that uses Autodiscover?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $autoMapping = $true
        }
        $totalFullAccessPermissionsExport = $fullAccessPermissions.count
        $FullAccessPermissionsCount = 0
        $currentFullAccessPermission = 0

        Write-Host 
        if($totalFullAccessPermissionsExport -eq 0) {
            
            $msg = "INFO: No FullAccess Permissions found in exported CSV file." 
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg  
              
        }
        else {
            $msg = "INFO: Importing FullAccess permissions into Office 365." 
            Write-Host $msg
            Log-Write -Message $msg

            Foreach($FullAccessPermission in $FullAccessPermissions){
            $currentFullAccessPermission += 1
            if($FullAccessPermission.DestinationPrimarySmtpAddress) {
                $targetMailbox = $FullAccessPermission.DestinationPrimarySmtpAddress
            }
            else {
                $targetMailbox = $FullAccessPermission.PrimarySmtpAddress
            }

            if($FullAccessPermission.DestinationDelegateAddress) {
                $delegate = $FullAccessPermission.DestinationDelegateAddress
            }
            else {
                $delegate = $FullAccessPermission.DelegateAddress
            }
            $recipientType = $FullAccessPermission.recipientType
            $displayName = $FullAccessPermission.DisplayName
            $delegateAccess = $FullAccessPermission.DelegateAccess
            $folderName = $FullAccessPermission.FolderName	 
            $folderType = $FullAccessPermission.FolderType   
            if ($autoMapping) {
                if($FullAccessPermission.AutoMapping -eq $false) {
                    $customAutoMapping = $false
                }
                else{
                    $customAutoMapping = $true
                }   
            } else {
                $customAutoMapping = $false
            }

            $mailboxList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”)

            if($FullAccessPermission.recipientType -in $mailboxList) {

                $msg = "INFO: Processing FullAccess permission $currentFullAccessPermission/$totalFullAccessPermissionsExport : TargetMailbox $targetMailbox Delegate $delegate AutoMapping $customAutoMapping."
                Write-Host $msg
                Log-Write -Message $msg
             
                #Verify if target mailbox exists    
                $recipient = check-O365Mailbox -mailbox $targetMailbox

                If ($recipient -eq $true) {

                    #Verify if delegate exists
                    $recipient = check-O365Mailbox -mailbox $delegate 
                    if($recipient -eq $false) {
                        $recipient = check-O365Group -group $delegate
                    }

                    If($recipient -eq $true) {
                        $result = Get-MailboxPermission -identity $targetMailbox -User $delegate 

                        if($result.AccessRights -eq "FullAccess") {
                            $msg = "      WARNING: FullAccess permission already exists in Office 365."
                            Write-Host -ForegroundColor Yellow $msg 
                            Log-Write -Message $msg                      
                        }
                        else {
                            try {
                                #If autoMapping -eq $true, it will check the autoMapping column in the CSV file
                                If($autoMapping -eq $true) {
                                    if($customAutoMapping -eq $false) {
                                        $autoMapping = $false
                                    }

                                    $result = Add-MailboxPermission -identity $targetMailbox -User $delegate -automapping $autoMapping -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
                                }
                                else {
                                    $result = Add-MailboxPermission -identity $targetMailbox -User $delegate -automapping $false -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
                                }

                                $msg = "      SUCCESS: FullAccess permission applied."
                                Write-Host -ForegroundColor Green $msg
                                Log-Write -Message $msg
                                $FullAccessPermissionsCount += 1   
                            }
                            catch {
                                $msg = "      ERROR: Failed to apply FullAccess permission."
                                Write-Host -ForegroundColor Red  $msg
                                Write-Host -ForegroundColor Red $_.Exception.Message
                                Log-Write -Message $msg
                                Log-Write -Message $_.Exception.Message
                            }                    
                        }
                    }
                    else {
                        $msg = "      ERROR: Delegate '$delegate' doest not exist in Office 365. FullAccess permission skipped."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg
                    }
        
                }
                else{
                    $msg =  "      ERROR: Target mailbox '$targetMailbox' doest not exist in Office 365."
                    Write-Host -ForegroundColor Red  $msg
                    Log-Write -Message $msg
                }  
            }
            else {
                Continue
            }
        }

            if($FullAccessPermissionsCount -ge 2) {
                $msg = "SUCCESS: $FullAccessPermissionsCount FullAccess permissions out of $totalFullAccessPermissionsExport have been applied to Office 365 mailboxes."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
            }elseif ($FullAccessPermissionsCount -eq 1) {
                $msg = "SUCCESS: 1 FullAccess permission out of $totalFullAccessPermissionsExport has been applied to Office 365 mailboxes."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
            }
        }
    }
    }
    else {
        $msg = "INFO: No FullAccess permissions found in Office 365." 
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg  
    }
}

#######################################################################################################################
#                                        IMPORT SENDONBEHALFOF PERMISSIONS INTO O365
#######################################################################################################################
Function Process-SendOnBehalfTo{

    $mailboxList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”)
    $groupList = @(“MailUniversalDistributionGroup”,“MailUniversalSecurityGroup”,“RoomList”)

    $msg = "INFO: Exporting SendOnBehalfTo permissions to CSV file." 
    Write-Host $msg
    Log-Write -Message $msg
    
    if($script:sendOnBehalfToPermissionsArray -ne $null) {    
        #Export sendOnBehalfToPermissionsArray to CSV file
        try {
            if($onlyPermissionsReport) {
                $script:sendOnBehalfToPermissionsArray | Export-Csv -Path $workingDir\O365SendOnBehalfToPermissionsReport.csv -NoTypeInformation -force
                $msg = "SUCCESS: CSV file '$workingDir\O365SendOnBehalfToPermissionsReport.csv' processed, exported and open."
            }
            else {
                $script:sendOnBehalfToPermissionsArray | Export-Csv -Path $workingDir\O365SendOnBehalfToPermissions.csv -NoTypeInformation -force
                $msg = "SUCCESS: CSV file '$workingDir\O365SendOnBehalfToPermissions.csv' processed, exported and open."
            }      

            Write-Host -ForegroundColor Green $msg
            Log-Write -Message $msg

            if ($script:sameEmailAddresses) {
                $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
            }
            elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "" -and $userDomain -eq $fullAccessDomain -and $onlyPermissionsReport -eq $false) {
                $msg = "         WARNING: The 'primarySmtpAddress' column of the opened CSV file has been updated with the new domain."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
                $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
            }
            elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""  -and $userDomain -ne $fullAccessDomain -and $onlyPermissionsReport -eq $false) {
                $msg = "         WARNING: The 'primarySmtpAddress' column of the opened CSV file has been updated with the new domain."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
                $msg = "         ACTION:  Populate the 'destinationGrantSendOnBehalfTo' column of the opened CSV file with the destination GrantSendOnBehalfTo userName."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg         
                $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
            }
            elseif(!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "" -and $onlyPermissionsReport -eq $false) {
                $msg = "         WARNING: Populate the 'destinationPrimarySmtpAddress' column of the opened CSV file with the destination user email."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg                
                $msg = "         WARNING: Populate the 'destinationGrantSendOnBehalfTo' column of the opened CSV file with the destination GrantSendOnBehalfTo userName."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg                
                $msg = "         ACTION:  Once you finish editing the CSV file, save it."
                Write-Host -ForegroundColor Yellow $msg
                Log-Write -Message $msg
            }
            elseif($onlyPermissionsReport -eq $false) {
            }
        }
        catch {
            if($onlyPermissionsReport) {
                $msg = "ERROR: Failed to import the CSV file '$workingDir\O365SendOnBehalfToPermissionsReport.csv'."
            }
            else {
                $msg = "ERROR: Failed to import the CSV file '$workingDir\O365SendOnBehalfToPermissions.csv'."
            }         
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        #Open the CSV file for editing
        if($onlyPermissionsReport) {
            Start-Process -FilePath $workingDir\O365SendOnBehalfToPermissionsReport.csv
        }
        else {
            Start-Process -FilePath $workingDir\O365SendOnBehalfToPermissions.csv
        }    

        #If the script must generate GSuite permissions report and also migrate them to O365
        if(!$onlyPermissionsReport) {
            $msg = "ACTION: If you have reviewed, edited and saved the CSV file then press any key to continue." 
            Write-Host $msg
            Log-Write -Message $msg
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

            #Re-import the edited CSV file
            Try{
                $sendOnBehalfToPermissions = @(Import-CSV "$workingDir\O365SendOnBehalfToPermissions.csv" | where-Object { $_.PSObject.Properties.Value -ne ""})
            }
            Catch [Exception] {
                $msg = "ERROR: Failed to import the CSV file '$workingDir\O365SendOnBehalfToPermissions.csv'. Please save and close the CSV file."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red $_.Exception.Message
                Log-Write -Message $msg
                Log-Write -Message $_.Exception.Message
                Exit
            }

            $totalSendOnBehalfToPermissionsExport = $sendOnBehalfToPermissions.count
            $sendOnBehalfToPermissionsCount = 0
            $currentSendOnBehalfToPermission = 0

            if($totalSendOnBehalfToPermissionsExport -eq 0) {
                Write-Host 
                $msg = "INFO: No SendOnBehalfTo permissions found in exported CSV file." 
                Write-Host -ForegroundColor Red  $msg
                Log-Write -Message $msg  
            }
            else {
                Write-Host 
                $msg = "INFO: Importing SendOnBehalfTo permissions into Office 365." 
                Write-Host $msg
                Log-Write -Message $msg

                Foreach($sendOnBehalfToPermission in $sendOnBehalfToPermissions){
                    $currentSendOnBehalfToPermission += 1
                    if($sendOnBehalfToPermission.DestinationPrimarySmtpAddress) {
                        $targetMailbox = $sendOnBehalfToPermission.DestinationPrimarySmtpAddress
                    }
                    else {
                        $targetMailbox = $sendOnBehalfToPermission.PrimarySmtpAddress
                    }

                    $recipientType = $sendOnBehalfToPermission.recipientType
                    $displayName = $sendOnBehalfToPermission.DisplayName
                    $grantSendOnBehalfTo = $sendOnBehalfToPermission.GrantSendOnBehalfTo

            
                    if($recipientType -in $mailboxList ) {   
                        $msg = "INFO: Processing SendOnBehalfTo permission $currentSendOnBehalfToPermission/$totalSendOnBehalfToPermissionsExport : TargetMailbox $targetMailbox SendOnBehalfTo $grantSendOnBehalfTo."
                        Write-Host $msg
                        Log-Write -Message $msg

                        #Verify if target mailbox exists    
                        $recipient = check-O365Mailbox -mailbox $targetMailbox
                    }
                    elseif($recipientType -in $groupList ) {  
                        $msg = "INFO: Processing SendOnBehalfTo permission $currentSendOnBehalfToPermission/$totalSendOnBehalfToPermissionsExport : TargetGroup $targetMailbox SendOnBehalfTo $grantSendOnBehalfTo."
                        Write-Host $msg
                        Log-Write -Message $msg 

                        #Verify if target mailbox exists    
                        $recipient = check-O365Group -group $targetMailbox
                    }
             
                    If ($recipient -eq $true) {

                        #Verify if sendOnBehalfTo user exists
                        $recipient = check-O365Mailbox -mailbox $grantSendOnBehalfTo 

                        If($recipient -eq $true) {

                            #Verify if sendOnBehalfTo permission already exists
                            try {
                                if($recipientType -in $mailboxList ) {    
                                    $result = Get-Mailbox -Identity $targetMailbox | Where-Object {$_.GrantSendOnBehalfTo -eq $grantSendOnBehalfTo} | select GrantSendOnBehalfTo
                                }                    
                                elseif($recipientType -in $groupList ) {  
                                    $result = Get-DistributionGroup -Identity $targetMailbox | Where-Object {$_.GrantSendOnBehalfTo -eq $grantSendOnBehalfTo} | select GrantSendOnBehalfTo
                                }
                            }
                            catch {
                                    $msg = "      ERROR: Failed to check SendOnBehalfTo permission."
                                    Write-Host -ForegroundColor Red  $msg
                                    Write-Host -ForegroundColor Red $_.Exception.Message
                                    Log-Write -Message $msg
                                    Log-Write -Message $_.Exception.Message
                            }
                            if($result -eq $null) {
                                try {
                                    if($recipientType -in $mailboxList ) {  
                                        $result = Set-Mailbox -identity $targetMailbox -GrantSendOnBehalfto @{Add=$grantSendOnBehalfTo} -ErrorAction Stop
                                    }
                                    elseif($recipientType -in $groupList ) {  
                                        $result = Set-DistributionGroup -identity $targetMailbox -GrantSendOnBehalfto @{Add=$grantSendOnBehalfTo} -ErrorAction Stop
                                    }
                                    $msg = "      SUCCESS: SendOnBehalfTo permission applied."
                                    Write-Host -ForegroundColor Green $msg
                                    Log-Write -Message $msg
                                    $sendOnBehalfToPermissionsCount += 1   
                                }
                                catch {
                                    $msg = "      ERROR: Failed to apply SendOnBehalfTo permission."
                                    Write-Host -ForegroundColor Red  $msg
                                    Write-Host -ForegroundColor Red $_.Exception.Message
                                    Log-Write -Message $msg
                                    Log-Write -Message $_.Exception.Message
                                }  
                            }
                            else {
                                $msg = "      WARNING: SendOnBehalfTo permission already exists in Office 365."
                                Write-Host -ForegroundColor Yellow $msg
                                Log-Write -Message $msg
                            }                          
                        }
                        else {
                            $msg = "      ERROR: SendOnBehalfTo '$grantSendOnBehalfTo' doest not exist in Office 365. SendOnBehalfTo permission skipped."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                        }
        
                    }
                    else{
                        if($recipientType -in $mailboxList ) {   
                            $msg =  "      ERROR: Target mailbox '$targetMailbox' doest not exist in Office 365."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                        }
                        elseif($recipientType -in $groupList ) {  
                            $msg =  "      ERROR: Target group '$targetMailbox' doest not exist in Office 365."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg
                        } 
                    }  
                }

                if($sendOnBehalfToPermissionsCount -ge 2) {
                    $msg = "SUCCESS: $sendOnBehalfToPermissionsCount FullAccess permissions out of $totalSendOnBehalfToPermissionsExport have been applied to Office 365 mailboxes."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg
                }
                elseif ($sendOnBehalfToPermissionsCount -eq 1) {
                    $msg = "SUCCESS: 1 FullAccess permission out of $totalSendOnBehalfToPermissionsExport has been applied to Office 365 mailboxes."
                    Write-Host -ForegroundColor Green $msg
                    Log-Write -Message $msg
                }
            }
        }
    }
    else {
        $msg = "INFO: No SendOnBehalfTo permissions found in Office 365." 
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg  
    }
}

#######################################################################################################################
#                                        IMPORT CALENDAR PERMISSIONS INTO O365
#######################################################################################################################
Function Process-FolderPermissions {
    
    $msg = "INFO: Exporting Calendar permissions to CSV file."  
    Write-Host $msg
    Log-Write -Message $msg
    
    if($script:folderPermissionsArray -ne $null) { 
    #Export calendarPermissionsArray to CSV file
    try {
        if($onlyPermissionsReport) {
            $script:folderPermissionsArray | Export-Csv -Path $workingDir\O365FolderPermissionsReport.csv -NoTypeInformation -force
            $msg = "SUCCESS: CSV file '$workingDir\O365FolderPermissionsReport.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg 
            Log-Write -Message $msg
        }
        else {
            $script:folderPermissionsArray | Export-Csv -Path $workingDir\O365CalendarPermissions.csv -NoTypeInformation -force
            $msg = "SUCCESS: CSV file '$workingDir\O365CalendarPermissions.csv' processed, exported and open."
            Write-Host -ForegroundColor Green $msg 
            Log-Write -Message $msg
        }        
        
        if ($script:sameEmailAddresses) {
            $msg = "         ACTION:  Please the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow $msg 
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne "" -and $userDomain -eq $calendarDomain -and $onlyPermissionsReport -eq $false) {
            $msg = "         WARNING: The 'primarySmtpAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         WARNING: The 'delegateAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and $script:sameUserName -and $script:destinationDomain -ne ""  -and $userDomain -ne $calendarDomain -and $onlyPermissionsReport -eq $false) {
            $msg = "         WARNING: The 'primarySmtpAddress' column of the opened CSV file has been updated with the new domain."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Populate the 'destinationDelegateAddress' column of the opened CSV file with the destination Calendar."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION:  Please review the opened CSV file and once you finish, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif(!$script:sameEmailAddresses -and !$script:sameUserName -and $script:destinationDomain -ne "" -and $onlyPermissionsReport -eq $false) {
            $msg = "         ACTION: Populate the 'destinationUser' column of the opened CSV file with the destination user email."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION: Populate the 'destinationDelegateAddress' column of the opened CSV file with the destination Calendar."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
            $msg = "         ACTION: Once you finish editing the CSV file, save it."
            Write-Host -ForegroundColor Yellow $msg
            Log-Write -Message $msg
        }
        elseif($onlyPermissionsReport -eq $false) {
        }
    }
    catch {
        if($onlyPermissionsReport) {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365FolderPermissionsReport.csv'."
        }
        else {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365CalendarPermissions.csv'."
        }         
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg
        Log-Write -Message $_.Exception.Message
        Exit
    }

    #Open the CSV file for editing
    if($onlyPermissionsReport) {
        Start-Process -FilePath $workingDir\O365FolderPermissionsReport.csv
    }
    else {
        Start-Process -FilePath $workingDir\O365CalendarPermissions.csv
    }    

    #If the script must generate GSuite permissions report and also migrate them to O365
    if(!$onlyPermissionsReport) {
        $msg = "ACTION: If you have reviewed, edited and saved the CSV file then press any key to continue." 
        Write-Host $msg
        Log-Write -Message $msg
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

        #Re-import the edited CSV file
        Try{
            $calendarPermissions = @(Import-CSV "$workingDir\O365CalendarPermissions.csv" | where-Object { $_.PSObject.Properties.Value -ne ""})
        }
        Catch [Exception] {
            $msg = "ERROR: Failed to import the CSV file '$workingDir\O365CalendarPermissions.csv'. Please save and close the CSV file."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        Write-Host
        $confirm = (Read-Host -prompt "Do you want to send an email with the published calendar URL to external users?  [Y]es or [N]o")
        if($confirm.ToLower() -eq "y") {
            $confirmExternalPermissions = $true
        }

        $totalCalendarPermissionsExport = $calendarPermissions.count
        $calendarPermissionsCount = 0
        $publishedCalendarUrlCount = 0
        $currentCalendarPermission = 0

        if($totalCalendarPermissionsExport -eq 0) {
            Write-Host 
            $msg = "INFO: No Calendar permissions found in exported CSV file."   
            Write-Host -ForegroundColor Red  $msg
            Log-Write -Message $msg
        }
        else {
            Write-Host 
            $msg = "INFO: Importing Calendar permissions into Office 365."
            Write-Host $msg
            Log-Write -Message $msg              

            Foreach($calendarPermission in $calendarPermissions){
                $currentCalendarPermission += 1 
                if($calendarPermission.DestinationPrimarySmtpAddress) {
                    $calendar = $calendarPermission.DestinationPrimarySmtpAddress
                }
                else {
                    $calendar = $calendarPermission.PrimarySmtpAddress
                }

                if($calendarPermission.DestinationDelegateAddress) {
                    $scope = $calendarPermission.DestinationDelegateAddress
                }
                else {
                    $scope = $calendarPermission.DelegateAddress
                }

                #$scopeType = $calendarPermission.scopeType
                $displayName = $calendarPermission.DisplayName
                $recipientType = $calendarPermission.RecipientType
                $delegateAccess = $calendarPermission.DelegateAccess
                $folderName = $calendarPermission.FolderName	 
                $folderType = $calendarPermission.FolderType  

                $mailboxList = @(“UserMailbox”,“SharedMailbox”,“RoomMailbox”,“EquipmentMailbox”,“TeamMailbox”,“GroupMailbox”)

                if($calendarPermission.recipientType -in $mailboxList) {
                    $msg = "INFO: Processing Calendar permission $currentCalendarPermission/$totalCalendarPermissionsExport : Calendar $calendar Delegate $scope Role $delegateAccess."
                    Write-Host $msg
                    Log-Write -Message $msg

                    #Verify if target mailbox exists
    
                    try {
                        $recipient = Get-Recipient -identity $calendar -ErrorAction SilentlyContinue
                    }
                    catch {
                       $msg = "      ERROR: Get-Recipient for $user failed."
                       Write-Host -ForegroundColor Red $msg
                       Write-Host -ForegroundColor Red $_.Exception.Message
                       Log-Write -Message $msg     
                       Log-Write -Message $_.Exception.Message       
                   }  
    
                    If ($recipient -ne $null) {
            
                        #Verify if scope exists
                        $isInternalScope = check-O365Mailbox -mailbox $scope
                        #Verify if mailboxFolderPermission exists
                        $folderPermission = Get-MailboxFolderPermission -Identity $calendar":\calendar" -user $scope  -ErrorAction SilentlyContinue
                        $scopeDomain = $scope.split("@")[1]

                        if($folderPermission) {
                                $msg = "      WARNING: Calendar permission already exists in Office 365."
                                Write-Host -ForegroundColor Yellow $msg
                                Log-Write -Message $msg
                        }
                        elseif(!$folderPermission -and $isInternalScope) {
            
                                $result = Add-MailboxFolderPermission -Identity $calendar":\calendar" -user $scope -AccessRights $delegateAccess -ErrorAction SilentlyContinue

                                if ($result) {
                                    $msg = "      SUCCESS: Calendar permission applied."
                                    Write-Host -ForegroundColor Green $msg
                                    Log-Write -Message $msg
                                    $calendarPermissionsCount += 1
                                }
                                else {
                                    $msg = "      ERROR: Calendar permission not applied."
                                    Write-Host -ForegroundColor Red $msg
                                    Log-Write -Message $msg
                                    
                                    $msg = "      CHECK: Get-MailboxFolderPermission -Identity $calendar"+":\calendar" + " -user $scope  -ErrorAction SilentlyContinue"
                                    Write-Host -ForegroundColor Yellow $msg
                                    Log-Write -Message $msg

                                    $calendarPermissionsCount += 1
                                }
                        }
                        elseif(!$folderPermission -and !$isInternalScope -and ($scopeDomain -eq $script:destinationDomain)) {
                            $msg =  "      ERROR: Scope '$scope' doest not exist in Office 365."
                            Write-Host -ForegroundColor Red  $msg
                            Log-Write -Message $msg

                            $msg = "      CHECK: Get-MailboxFolderPermission -Identity $calendar"+":\calendar" + " -user $scope  -ErrorAction SilentlyContinue"
                            Write-Host -ForegroundColor Yellow $msg
                            Log-Write -Message $msg
                        }
                        elseif(!$folderPermission -and !$isInternalScope -and ($scopeDomain -ne $script:destinationDomain)) {
                            $msg = "      WARNING: User '$scope' does not exist in Office 365, is an external user." 
                            Write-Host -ForegroundColor Yellow $msg
                            Log-Write -Message $msg

                            if($confirmExternalPermissions) {
                        
                                $publishedCalendar = Get-MailboxCalendarFolder -Identity "$($recipient.Identity):\calendar" 
                                if (!$publishedCalendar.publishEnabled) {
                                    $result = Set-MailboxCalendarFolder -Identity "$($recipient.Identity):\calendar" -DetailLevel AvailabilityOnly -PublishEnabled $true -ErrorAction SilentlyContinue

                                    if ($result) {
                                        $msg = "      SUCCESS: Calendar Sharing URL published."
                                        Write-Host -ForegroundColor Green $msg
                                        Log-Write -Message $msg

                                        $publishedCalendar = Get-MailboxCalendarFolder -Identity "$($recipient.Identity):\calendar" -ErrorAction SilentlyContinue
                                    }
                                    else {
                                        $msg = "      ERROR:  Calendar Sharing URL not published."
                                        Write-Host -ForegroundColor Red $msg
                                        Log-Write -Message $msg
                                        Continue
                                    }
                                }                        

                                #published calendar HTML URL
                                $htmlPublishedCalendarUrl=$publishedCalendar.PublishedCalendarUrl
                                #published calendar ICS URL
                                $icsPublishedCalendarURL = $publishedCalendar.PublishedICalUrl 
                                $smtpServer = "smtp.office365.com"
                                $smtpCreds = $script:destinationO365Creds
                                $emailTo = $scope
                                $emailFrom = $smtpCreds.Username
                                $FirstName = $recipient.FirstName
                                $LastName = $recipient.LastName
                                $subject = "You're invited to share this calendar"
                                $body = ""
                                $body += "<h2>I'd like to share my calendar with you </h2><br>"
                                $body += "$FirstName $LastName ($calendar) would like to share an Outlook calendar with you. <br><br>"
                                $body += "You'll be able to see the availability information of events on <a href="+$htmlPublishedCalendarUrl+">this calendar</a>. <br><br>"
                                $body += "To import the calendar to your Outlook calendar, this is the <a href="+$icsPublishedCalendarURL+">ICS file</a>."

                                try {
                                    $result = Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpServer -Port 587 -Credential $smtpCreds -UseSsl -ErrorAction SilentlyContinue
                        
                                    if ($error[0].ToString() -match "Spam abuse detected from IP range.") { 
                                        #5.7.501 Access denied, spam abuse detected. The sending account has been banned due to detected spam activity. 
                                        #For details, see Fix email delivery issues for error code 451 5.7.500-699 (ASxxx) in Office 365.
                                        #https://support.office.com/en-us/article/fix-email-delivery-issues-for-error-code-451-4-7-500-699-asxxx-in-office-365-51356082-9fef-4639-a18a-fc7c5beae0c8 
                                        $msg = "      ERROR: Failed to send email to user '$emailTo'. Access denied, spam abuse detected. The sending account has been banned. "
                                        Write-Host -ForegroundColor Red  $msg
                                        Log-Write -Message $msg
                                    }
                                    else {
                                        $msg = "      SUCCESS: Email with $FirstName's calendar URL sent to external user '$emailTo'"
                                        Write-Host -ForegroundColor Green $msg
                                        Log-Write -Message $msg 
                                        $publishedCalendarUrlCount += 1   
                                   }
  
                                }
                                catch {
                                    $msg = "      ERROR: Failed to send email to user '$emailTo'."
                                    Write-Host -ForegroundColor Red  $msg
                                    Write-Host -ForegroundColor Red $_.Exception.Message
                                    Log-Write -Message $msg
                                    Log-Write -Message $_.Exception.Message
                                }

                            }
                        }
        
                    }
                    else{
                        $msg =  "      ERROR: Target mailbox '$calendar' doest not exist in Office 365."
                        Write-Host -ForegroundColor Red  $msg
                        Log-Write -Message $msg
                    }    
                }
                else {
                    Continue
                }
            }

            if($calendarPermissionsCount -ge 2) {
                $msg = "SUCCESS: $calendarPermissionsCount Calendar permissions out of $totalCalendarPermissionsExport have been applied to Office 365 mailboxes."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
            }
            elseif ($calendarPermissionsCount -eq 1) {
                $msg = "SUCCESS: 1 Calendar permission out of $totalCalendarPermissionsExport has been applied to Office 365 mailboxes."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
            }

            if($publishedCalendarUrlCount -ge 2) {
                $msg = "SUCCESS: $publishedCalendarUrlCount published Calendar URLs have been sent to external users."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
            }elseif ($publishedCalendarUrlCount -eq 1) {
                $msg = "SUCCESS: 1 published Calendar URL has been sent to an external user."
                Write-Host -ForegroundColor Green $msg
                Log-Write -Message $msg
            }
        }
    }
    }
    else {
        $msg = "INFO: No Calendar permissions found in Office 365." 
        Write-Host -ForegroundColor Red  $msg
        Log-Write -Message $msg  
    }
}

#######################################################################################################################
#                                               MAIN MENU
#######################################################################################################################
# Function to display the main menu
Function Menu-O365T2T {

    #Main menu
    do {
        Write-Host 
        $confirm = (Read-Host -prompt "What do you want to do? `
        1. Generate Office 365 permissions [r]eports `
        2. Generate user [b]atches based on FullAccess permissions `
        3. Migrate distribution [g]roups and mail-enabled security groups `
        4. [M]igrate mailbox, calendar and group permissions `
        5. [E]xit

        [R]eports, [B]atches, [G]roups, [M]igrate or [E]xit")

        if($confirm.ToLower() -eq "r") {
            $onlyPermissionsReport=$true
            Write-Host
            $msg = "INFO: This script is going to only export mailbox, folder and/or groups permissions from source Office 365 to CSV files." 
            Write-Host $msg
            Log-Write -Message $msg
            Write-Host 
            if($script:sourceO365Session.State -ne 'opened') {       
                $script:sourceO365Session = Connect-SourceExchangeOnline
            }
            else{
                $msg = "INFO: Already connected to source Office 365 Remote PowerShell."
                Write-Host $msg
                Log-Write -Message $msg
            }
            Write-Host
        }
        elseif ($confirm.ToLower() -eq "b") {
            $onlyPermissionsReport=$false
            Write-Host
            $msg = "INFO: This script is going to export all FullAccess mailbox permissions from source Office 365"
            Write-Host $msg
            Log-Write -Message $msg
            $msg = "      and generate user batches based on these exported FullAccess permissions." 
            Write-Host $msg
            Log-Write -Message $msg
            Write-Host    
            if($script:sourceO365Session.State -ne 'opened') {       
                $script:sourceO365Session = Connect-SourceExchangeOnline
            }
            else{
                $msg = "INFO: Already connected to source Office 365 Remote PowerShell."
                Write-Host $msg
                Log-Write -Message $msg
            }
            Write-Host
            $createUserBatches = $true
        }
        elseif ($confirm.ToLower() -eq "m") {
            $onlyPermissionsReport=$false
            Write-Host
            $msg = "INFO: This script is going to export mailbox and/or folder permission from source Office 365 to CSV files"
            Write-Host $msg
            Log-Write -Message $msg
            $msg = "      and import them into destination Office 365." 
            Write-Host $msg
            Log-Write -Message $msg
            Write-Host    
            if($script:sourceO365Session.State -ne 'opened') {       
                $script:sourceO365Session = Connect-SourceExchangeOnline
            }
            else{
                $msg = "INFO: Already connected to source Office 365 Remote PowerShell."
                Write-Host $msg
                Log-Write -Message $msg
            }
            if($script:destinationO365Session.State -ne 'opened') {       
                $script:destinationO365Session = Connect-DestinationExchangeOnline
            }    
            else{
                $msg = "INFO: Already connected to destination Office 365 Remote PowerShell."
                Write-Host $msg
                Log-Write -Message $msg
            }        
            Write-Host
            query-EmailAddressMapping
            Write-Host
        }
        elseif ($confirm.ToLower() -eq "g") {
            $onlyPermissionsReport=$false
            Write-Host
            $msg = "INFO: This script is going to export distribution groups from source Office 365 to CSV files"
            Write-Host $msg
            Log-Write -Message $msg
            $msg = "      and import them into destination Office 365." 
            Write-Host $msg
            Log-Write -Message $msg
            Write-Host    
            $msg = "WARNING: The destination distribution groups will be created with the same source names."
            Write-Host -ForegroundColor yellow $msg
            Log-Write -Message $msg
            Write-Host
            if($script:sourceO365Session.State -ne 'opened') {       
                $script:sourceO365Session = Connect-SourceExchangeOnline
            }
            else{
                $msg = "INFO: Already connected to source Office 365 Remote PowerShell."
                Write-Host $msg
                Log-Write -Message $msg
            }
            if($script:destinationO365Session.State -ne 'opened') {       
                $script:destinationO365Session = Connect-DestinationExchangeOnline
            } 
            else{
                $msg = "INFO: Already connected to destination Office 365 Remote PowerShell."
                Write-Host $msg
                Log-Write -Message $msg
            }  
    
            $createDistributionGroups = $true

            Write-Host
            query-EmailAddressMapping
            Write-Host
            Export-O365Groups
            Write-Host
            Return 1
        }
        elseif ($confirm.ToLower() -eq "e") {
            Return $null
        }

    } while(($confirm.ToLower() -ne "r") -and ($confirm.ToLower() -ne "m") -and ($confirm.ToLower() -ne "b") -and ($confirm.ToLower() -ne "g") -and ($confirm.ToLower() -ne "e"))

    # Skip the users that do not exist in destination Office 365
    $skipNonExistingUser = $false
    do {
        $confirm = (Read-Host -prompt "Do you want to skip the users that do not exist in destination Office 365?  [Y]es or [N]o")

        if($confirm.ToLower() -eq "y") {
            $skipNonExistingUser=$true
            if(!$script:destinationO365Session) {
                if($script:destinationO365Session.State -ne 'opened') {       
                    $script:destinationO365Session = Connect-DestinationExchangeOnline
                } 
                else{
                    $msg = "INFO: Already connected to destination Office 365 Remote PowerShell."
                    Write-Host $msg
                    Log-Write -Message $msg
                }  
                Write-Host
                query-EmailAddressMapping
                Write-Host
            }
        }

    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
    
    # Import a CSV file with the users to process
    $readCSVFile = $false
    do {
        $confirm = (Read-Host -prompt "Do you want to import a CSV file with the users you want to process?  [Y]es or [N]o")

        if($confirm.ToLower() -eq "y") {
            $readCSVFile = $true
        
            Write-Host -ForegroundColor yellow "ACTION: Select the CSV file to import file (Press cancel to create one)"
        
            Get-FileName $workingDir
        }

    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    #For user batch creation, FullAccess permissions are taken into consideration.
    if($createUserBatches -eq $true)  {        
        $processSendAs = $false
        $processFullAccess = $true
        $processFolders = $false
        $processSendOnBehalfTo = $false
        $processGroups = $false

        # Export only FullAccess permissions from Office 365 for user batch creation
        $result = Export-O365Permissions -skipNonExistingUser $skipNonExistingUser `
                                         -processSendAs $processSendAs `
                                         -processFullAccess  $processFullAccess `
                                         -processFolders $processFolders `
                                         -processSendOnBehalfTo $processSendOnBehalfTo `
                                         -processGroups $processGroups `
                                         -userBatches $createUserBatches `
                                         -readCSVfile $readCSVfile 
    }
    # For everything else, which permissions should be included in the processing
    else {

        $processSendAs = $false
        $processFullAccess = $false
        $processFolders = $false
        $processSendOnBehalfTo = $false
        $processOnlyCalendars = $false
        $processGroups = $false
 
        Write-Host

        # SendAs
        do {
            if ($onlyPermissionsReport) {
                $confirm = (Read-Host -prompt "Do you want to generate source Office 365 SendAs mailbox permissions report?  [Y]es or [N]o")    
            }
            else {
                $confirm = (Read-Host -prompt "Do you want to migrate SendAs permissions from source Office 365 to destination Office 365?  [Y]es or [N]o")
            }
            if($confirm.ToLower() -eq "y") {
                $processSendAs = $true
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))


        # FullAccess
        do {
            if ($onlyPermissionsReport) {
                $confirm = (Read-Host -prompt "Do you want to generate source Office 365 FullAccess mailbox permissions report?  [Y]es or [N]o")    
            }
            else{
                $confirm = (Read-Host -prompt "Do you want to migrate FullAccess permissions from source Office 365 to destination Office 365?  [Y]es or [N]o")
            }
            
            if($confirm.ToLower() -eq "y") {
                $processFullAccess = $true
            }

        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        # SendOnBehalfOf
        do {
            if ($onlyPermissionsReport) {
                $confirm = (Read-Host -prompt "Do you want to generate source Office 365 SendOnBehalfOf mailbox permissions report?  [Y]es or [N]o")   
            }
            else{
                $confirm = (Read-Host -prompt "Do you want to migrate SendOnBehalfOf permissions from source Office 365 to destination Office 365?  [Y]es or [N]o")
            }

            if($confirm.ToLower() -eq "y") {
                $processSendOnBehalfTo = $true
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))


        # Calendar
        do {
            if ($onlyPermissionsReport) {
                $confirm = (Read-Host -prompt "Do you want to generate source Office 365 mailbox Folder permissions report?  [Y]es or [N]o")   
                if($confirm.ToLower() -eq "y") {
                     do {
                        $confirmCalendars = (Read-Host -prompt "Do you want to only include Calendar permissions in the Folder permissions report?  [Y]es or [N]o")  
                        if($confirmCalendars.ToLower() -eq "y") {
                            $processOnlyCalendars = $true
                        }  
                    } while(($confirmCalendars.ToLower() -ne "y") -and ($confirmCalendars.ToLower() -ne "n"))
                }
            }
            else{
                # All Folder permissions are migrated via MW. By applying MustMigrateAllPermissions=1 also non-system folder permissions are migrated : Migrate all folder permissions, not just the system folder ones.
                $confirm = (Read-Host -prompt "Do you want to migrate Calendar permissions from source Office 365 to destination Office 365?  [Y]es or [N]o")
                $processOnlyCalendars = $true
            }

            if($confirm.ToLower() -eq "y") {
                $processFolders = $true
            } 

        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
        
        # Distribution groups and mail enabled security groups
        do {
            if ($onlyPermissionsReport) {
                $confirm = (Read-Host -prompt "Do you want to generate source distribution lists and mail-enabled security group SendAs and SendOnBehalfOf permissions report?  [Y]es or [N]o")    
            }
            else {
                $confirm = (Read-Host -prompt "Do you want to migrate distribution list and mail-enabled security group SendAs and SendOnBehalfOf permissions from source Office 365 to destination Office 365?  [Y]es or [N]o")
            }

            if($confirm.ToLower() -eq "y") {
                $processGroups = $true
            }
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        # Export permissions from Office 365 based on the previous parameters
        if(($processSendAs -eq $true) -or ($processFullAccess -eq $true) -or ($processFolders -eq $true) -or ($processSendOnBehalfTo -eq $true) -or ($processOnlyCalendars -eq $true) -or ($processGroups -eq $true)) {
                    
            $result = Export-O365Permissions -skipNonExistingUser $skipNonExistingUser `
                                             -processSendAs $processSendAs `
                                             -processFullAccess  $processFullAccess `
                                             -processFolders $processFolders `
                                             -processSendOnBehalfTo $processSendOnBehalfTo `
                                             -processGroups $processGroups `
                                             -userBatches $createUserBatches `
                                             -readCSVfile $readCSVfile 
        
            # Import into destination Office 365 permissions previously exported from source Office 365
            if($processSendAs) {
                Write-Host
                Process-SendAsPermissions 
                Start-Sleep -s 5
            }
            if($processFullAccess) {
                Write-Host
                Process-FullAccessPermissions
                Start-Sleep -s 5
            }
            if ($processSendOnBehalfTo){
                Write-Host
                Process-SendOnBehalfTo
                Start-Sleep -s 5
            }
            if($processFolders) {
                Write-Host
                Process-FolderPermissions
                Start-Sleep -s 5
            }
        }
    }


    Return 1

}

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################

## Initiate Parameters

#Working Directory
$workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Migrate-PermissionsO365toO365.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($logFile)'."
Write-Host -ForegroundColor Yellow "         All CSV files will be in folder '$($workingDir)'."
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg

#To only have source email addresses when only generating reports
$script:sameEmailAddresses = $true
$script:differentDomain = $false
$script:sameUserName = $true
$script:destinationDomain = "" 
$createUserBatches = $false
$createDistributionGroups = $false


# keep looping until specified to exit
do {
    $action = Menu-O365T2T
	if($action -ne $null) {
			$action = Menu-O365T2T
	}
	else {
        ##END SCRIPT 
	    Write-Host

        $msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
        Log-Write -Message $msg

        if($sourceO365Session) {

            try {
                Write-Host "INFO: Opening directory $workingDir where you will find all the generated CSV files."
                Invoke-Item $workingDir
                Write-Host
            }
            catch{
                $msg = "ERROR: Failed to open directory '$workingDir'. Script will abort."
                Write-Host -ForegroundColor Red $msg
                Exit
            }

            Remove-PSSession $sourceO365Session
            if($destinationO365Session) {
                Remove-PSSession $destinationO365Session
            }
        }

        Exit
	}
}
while($true)



