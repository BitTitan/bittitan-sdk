# Copyright © MigrationWiz 2011.  All rights reserved.

$smtpServerHost = "smtp.mail.microsoftonline.com"
$smtpServerPort = 587
$smtpServerSsl  = $true

$emailFrom      = '"IT Admin" <admin@mso.bittitan.com>'
$emailSubject   = "IMPORTANT: Your mailbox is about to be migrated"
$emailFile      = ".\EmailUsers.htm"

$importFile     = "Users.csv"


&{
    $write = "Enter the user name and password to access " + $smtpServerHost + ":" + $smtpServerPort
    Write-Host $write -foreground "yellow"
    Write-Host ""

    $cred = Get-Credential
    Write-Host ""

    $users = import-csv $importFile | select *
    foreach($user in $users)
    {
        $location       = $user.'Mailbox Location'
        $size           = $user.'Mailbox Size'
        $emailAddress   = $user.'Email Address'
        $password       = $user.'Password'
        $firstName      = $user.'First Name'
        $lastName       = $user.'Last Name'
        $displayName    = $user.'Display Name'
        $title          = $user.'Job Title'
        $department     = $user.'Department'
        $officeNumber   = $user.'Office Number'
        $officePhone    = $user.'Office Phone'
        $mobile         = $user.'Mobile Phone'
        $fax            = $user.'Fax'
        $address        = $user.'Address'
        $city           = $user.'City'
        $state          = $user.'State or Province'
        $postalCode     = $user.'ZIP or Postal Code'
        $country        = $user.'Country or Region'
        
        $write = "Emailing account " + $emailAddress + " ..."
        Write-Host $write
        
        $body = [string]::join([environment]::newline, (Get-Content -path $emailFile))
        
        $body = $body.Replace('[EmailAddress]', $emailAddress)
        $body = $body.Replace('[Password]', $password)
        $body = $body.Replace('[FirstName]', $firstName)
        $body = $body.Replace('[LastName]', $lastName)
        
        $mail = New-Object System.Net.Mail.MailMessage 
        $mail.From = $emailFrom 
        $mail.To.Add($emailAddress) 
        $mail.Subject = $emailSubject
        $mail.IsBodyHtml = $true
        $mail.Body = $body

        $smtp = New-Object System.Net.Mail.SmtpClient
        $smtp.Host = $smtpServerHost
        $smtp.Port = $smtpServerPort
        $smtp.EnableSsl = $smtpServerSsl
        $smtp.UseDefaultCredentials = $false
        $smtp.Credentials = New-Object System.Net.NetworkCredential($cred.UserName, $cred.GetNetworkCredential().Password)
        $smtp.Send($mail)
        
        Write-Host ""
    }
}
trap
{
    break;
}   