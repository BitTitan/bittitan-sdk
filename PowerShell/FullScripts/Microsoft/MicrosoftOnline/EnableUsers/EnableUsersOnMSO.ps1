# Copyright © MigrationWiz 2011.  All rights reserved.

$importFile = "Users.csv"

&{
    $write = "Enter the admin user name and password to Microsoft Online"
    Write-Host $write -foreground "yellow"
    Write-Host ""

    $cred = Get-Credential
    Write-Host ""

    $subscriptions = Get-MSOnlineSubscription -Credential $cred
    foreach($subscription in $subscriptions)
    {
        $write = "Subscription ID: " + $subscription.SubscriptionId + "`r`n"
        foreach($type in $subscription.SubscriptionServiceTypes)
        {
            $write += "                 " + $type + "`r`n"
        }
        Write-Host $write
    }

    Write-Host 'Enter the Subscription ID to assign:' -foreground "yellow"
    $subscriptionId = Read-Host
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
        
        $write = "Enabling account " + $emailAddress + " ..."
        Write-Host $write
        
        $result = Enable-MSOnlineUser -Credential $cred -Identity $emailAddress -UserLocation $location -MailboxQuotaSize $size -Password $password -SubscriptionIDs $subscriptionId
        
        Write-Host ""
    }
}
trap
{
    break;
}   