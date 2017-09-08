# Copyright © MigrationWiz 2011.  All rights reserved.

$importFile = "Users.csv"

&{
    $write = "Enter the admin user name and password to Microsoft Online"
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
        
        $write = "Updating account " + $emailAddress + " ..."
        Write-Host $write

        $result = Set-MSOnlineUser -Credential $cred -Identity $emailAddress -DisplayName $displayName -FirstName $firstName -LastName $lastName -JobTitle $title -Department $department -OfficeNumber $officeNumber -OfficePhone $officePhone -MobilePhone $mobile -FaxNumber $fax

        Write-Host ""
    }
}
trap
{
    break;
}    