<#
.NOTES
    Company:          BitTitan, Inc.
    Title:            Export-mail-enabled-public-folder-email-addresses.ps1
    Author:           Support@BitTitan.com
    Version:          1.01
    Date:             March 6, 2019
    Disclaimer:       This script is provided 'AS IS'. No warranty is provided either expressed or implied
    Copyright:        Copyright © 2019 BitTitan. All rights reserved
.SYNOPSIS
    This script discovers the mail enabled public folders at the source so their addresses can be addeded to the migrated mail enabled pulic folders at the destination.
.DESCRIPTION
    This script creates a CSV listing all mail enabled public folder SMTP addresses from the source including the folder path for which they belong. 
    Addresses in the EmailAddresses column with a uppercase SMTP: depict the current primary SMTP address for the source mail enabled folder. 
    Addresses with a lowercase smtp: depict a current alias address assigned to the source mail enabled folder. 
#>

# declare the results array
$resultsarray = @()

# Determine if the identity can be used to retrieve public folders
function CanRetrievePublicFoldersWithIdentity
{
    if ((Get-PublicFolder -Identity $args[0].Identity -ErrorAction SilentlyContinue) -ne $null) { return $true }
    else { return $false }
}

# Load public folders by recipient ID
function LoadPublicFoldersByRecipientId
{
    $publicFolders = Get-PublicFolder -Recurse -ResultSize unlimited | Where {$_.MailEnabled -eq "True"}
    $pfDictionary = @{}
    foreach ($publicFolder in $publicFolders) { $pfDictionary.Add($publicFolder.MailRecipientGuid, $publicFolder) }
    return $pfDictionary
}

# Retrieve all mail enabled public folders
[array]$mailpub = Get-MailPublicFolder -ResultSize Unlimited
if ($mailpub.length -gt 0)
{
    $useIdentity = CanRetrievePublicFoldersWithIdentity $mailpub[0]
    
    # When the identity cannot be used we have to pre-load all public folders
    if ($useIdentity -eq $false)
    {    
        $pfDictionary = LoadPublicFoldersByRecipientId
    }
}

# Go through all folders and extract the relevant data
foreach ($folder in $mailpub)
{
    # set the fields
    $email = $folder.primarysmtpaddress.local + "@" + $folder.primarysmtpaddress.domain
    if ($useIdentity -eq $true)
    {
        $pubfolder  = Get-PublicFolder -Identity $folder.Identity
    }

    else
    {
	Write-Output "Looking up folder $($folder.Guid)"
        $pubfolder  = $pfDictionary[$folder.Guid]
    }

    Write-Output "Processing folder $($pubfolder.name) with parent $($pubfolder.parentpath)"
  
    # set the folder path and trim any leading or trailing whitespaces
    if ($pubfolder.parentpath -eq "\")
    {
	$folderpath = "\" + ($pubfolder.name.trim())
    }
  
    else
    {
        $folderpath = ($pubfolder.parentpath.trim()) + "\" + ($pubfolder.name.trim())
    }

    # go through all email addresses
    foreach ($alternateemail in $folder.emailaddresses)
    {
	$alternateString = $alternateEmail.ToString()
	
	# skip non-SMTP addresses
	if ($alternateString.StartsWith("SMTP:", "CurrentCultureIgnoreCase") -ne $true)
        {
	    continue
	};
	
	# set the fields
	$email = $alternateString.SubString("SMTP:".Length)
	$proxy = $alternateString
	
	# create the object to add to the result set
	$pubObject = new-object PSObject
	$pubObject | add-member -membertype NoteProperty -name "SmtpAddress" -Value $email
	$pubObject | add-member -membertype NoteProperty -name "EmailAddresses" -Value $proxy
	$pubObject | add-member -membertype NoteProperty -name "FolderPath" -Value $folderpath
	
	# add the object to the result set
	$resultsarray += $pubObject
    }
}

if ($resultsarray.length -gt 0)
{
    # Finally we want export our results to a CSV
    $path = "mail-enabled-public-folder-email-addresses.csv"
    $resultsarray | export-csv -Path $path

    # Output that the results were written
    "Success: Output the email addresses to CSV file: " + $path
}

else
{
    Write-Warning "No mail enabled public folders found"
}

# output any errors to a text file for customer to review if needed
$location = Get-Location

Write-Host "`nExport is complete. Any errors generated have been saved to ExportErrorlog.txt in $location.`n" -ForegroundColor Green

$error | Out-File .\ExportErrorlog.txt
