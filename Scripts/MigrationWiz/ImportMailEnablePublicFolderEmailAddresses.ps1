## import-mail-enabled-public-folder-email-addresses.ps1 ##
###########################################################

# check if a file has been passed in 
$file = $Args[0]
if ($file -eq $null) {
  $file = ".\mail-enabled-public-folder-email-addresses.csv"
}

# make sure the file exists to import
if (test-path $file) {
  # output the file that is being imported
  "Importing from " + $file
} else {
  Write-Warning "Cannot find file to import"
  exit
}

# import the CSV file which contains all of the mail enabled public folder email addresses
$importValues = import-csv $file

# determine the unique folders
$folders = @()
foreach ($value in $importValues) {
  $folders += $value.FolderPath
}
$folders = $folders | select -unique

# go through each folder and attempt to add the email addresses
foreach ($folder in $folders) {

  # get the existing mail enabled public folder
  $publicFolder = Get-MailPublicFolder $folder
  
  # continue if the folder does not exist
  if ($publicFolder -eq $null) {
	Write-Warning {"Could not find mail enabled public folder " + $folder + " to add email addresses.  Skipping processing."}
	continue
  };
  
  # go through all imported values looking for email addresses
  foreach ($value in $importValues)
  {
    # do not import the primary email address
    if ($value.IsPrimary -eq $true) {
	  continue
	};
  
    # only add the email address if the folders match
	if ($value.FolderPath -eq $folder) {
		if ($publicFolder.EmailAddresses -notcontains $value.SmtpAddress) {
			$publicFolder.EmailAddresses += $value.SmtpAddress
		}
	}
  }
  
  # set the email addresses
  Set-MailPublicFolder -Identity $publicFolder.Identity -EmailAddresses $publicFolder.EmailAddresses
}