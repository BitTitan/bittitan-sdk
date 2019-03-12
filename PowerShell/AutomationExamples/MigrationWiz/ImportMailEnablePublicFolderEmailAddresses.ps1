<#
.NOTES
    Company:          BitTitan, Inc.
    Title:            Import-mail-enabled-public-folder-email-addresses.ps1
    Author:           Support@BitTitan.com
    Version:          1.01
    Date:             March 6, 2019
    Disclaimer:       This script is provided 'AS IS'. No warranty is provided either expressed or implied
    Copyright:        Copyright © 2019 BitTitan. All rights reserved
.SYNOPSIS
    This script adds the smtp addresses currently assigned to the source mail enabled public folders to the newly migrated mail enabled public folders while also setting the primary SMTP address.
.DESCRIPTION
    This script is designed to be ran once for newly migrated mail enabled public folders at the destination.
    It uses the CSV created from the export script and adds the smtp addresses in the EmailAddresses column to the corresponding mail enabled public folders at the destination.
    Any address in the EmailAddresses column that begins with an uppercase SMTP: will become the primary SMTP address of the mail enabled folder in the destination. 
#>

# check if a file has been passed in 
$file = $Args[0]
if ($file -eq $null)
{
    $file = ".\mail-enabled-public-folder-email-addresses.csv"
}

# make sure the file exists to import
if (test-path $file)
{
    # output the file that is being imported
    "Importing from " + $file
}

else
{
    Write-Warning "Cannot find file to import"
    exit
}

# import the CSV file which contains all of the mail enabled public folder email addresses
$importValues = import-csv $file

# determine the unique folders
$folders = @()
foreach ($value in $importValues)
{
    $folders += $value.FolderPath
}

$folders = $folders | select -unique

# go through each folder and attempt to add the email addresses
foreach ($folder in $folders)
{
    # replace / with -
    $folder = $folder.Replace('/', '_')

    # get the existing mail enabled public folder
    $publicFolder = Get-MailPublicFolder $folder
  
    # continue if the folder does not exist
    if ($publicFolder -eq $null)
    {
        Write-Warning ("Could not find mail enabled public folder " + $folder + " to add email addresses.  Skipping processing.") -WarningAction Inquire
        continue
    };
  
    # go through all imported values looking for email addresses
    foreach ($value in $importValues)
    {
        if ($value.FolderPath -eq $folder)
	{
	    $publicFolder.EmailAddresses += $value.EmailAddresses
        }
    }
  
  
  # set the email addresses
  Set-MailPublicFolder -Identity $publicFolder.Identity -EmailAddresses $publicFolder.EmailAddresses -EmailAddressPolicyEnabled $false
}

# output any errors to a text file for customer to review if needed
$location = Get-Location

Write-Host "`nImport is complete. Any errors generated have been saved to ImportErrorlog.txt in $location.`n" -ForegroundColor Green

$error | Out-File .\ImportErrorlog.txt
