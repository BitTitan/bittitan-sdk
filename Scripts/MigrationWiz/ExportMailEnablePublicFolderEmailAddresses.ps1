## export-mail-enabled-public-folder-email-addresses.ps1 ##
###########################################################

# declare the results array
$resultsarray = @()

# retrieve all mail enabled public folders
$mailpub = Get-MailPublicFolder -ResultSize unlimited
foreach ($folder in $mailpub) {
  # set the fields
  $email      = $folder.primarysmtpaddress.local + "@" + $folder.primarysmtpaddress.domain
  $pubfolder  = Get-PublicFolder -Identity $folder.identity
  
  # set the folder path
  if ($pubfolder.parentpath -eq "\") {
	$folderpath = "\" + $pubfolder.name
  } else {
	$folderpath = $pubfolder.parentpath + "\" + $pubfolder.name
  }

  # go through all email addresses
  foreach ($alternateemail in $folder.emailaddresses) {
	# skip non-SMTP addresses
	if ($alternateemail.PrefixString -ne "SMTP") {
		continue 
	};
	
	# set the fields
	$email = $alternateemail.SmtpAddress
	$prefix = $alternateemail.PrefixString
	$proxy = $alternateemail.ProxyAddressString
	
	# create the object to add to the result set
	$pubObject = new-object PSObject
	$pubObject | add-member -membertype NoteProperty -name "SmtpAddress" -Value $email
	$pubObject | add-member -membertype NoteProperty -name "ProxyAddressString" -Value $proxy
	$pubObject | add-member -membertype NoteProperty -name "Prefix" -Value $prefix
	$pubObject | add-member -membertype NoteProperty -name "IsPrimary" -Value $false
	$pubObject | add-member -membertype NoteProperty -name "FolderPath" -Value $folderpath
	
	# add the object to the result set
	$resultsarray += $pubObject
  }
}

if ($resultsarray.length -gt 0) {
  # Finally we want export our results to a CSV
  $path = "mail-enabled-public-folder-email-addresses.csv"
  $resultsarray | export-csv -Path $path

  # Output that the results were written
  "Success: Output the email addresses to CSV file: " + $path
} else {
  Write-Warning "No mail enabled public folders found"
}