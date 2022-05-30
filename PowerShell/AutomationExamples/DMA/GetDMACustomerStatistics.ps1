<#
.NOTES
    Company:		BitTitan, Inc.
    Title:			GetDMACustomerStatistics.PS1
    Author:			SUPPORT@BITTITAN.COM
    Requirements: 
    
    Version:		1.01
    Date:			April 6, 2017

    Windows Version:	WINDOWS 10 ENTERPRISE

    Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

    Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
    
.SYNOPSIS
    Generates a full csv report of DMA statistics for a given customer.

.DESCRIPTION 	
    This script retrieves all DMA statistics data for a given customer and generates a report. 

.INPUTS
    -[ManagementProxy.ManagementService.Ticket] Ticket, the ticket for authentication.
    -[guid] CustomerId, the id of the customer to report.
    -[string] Csv, the csv output path.
    -[int] PageSize, the batch size of every retrieve request.
    -[string] Env, the context to work with. Valid options : BT, China.

.EXAMPLE
    .\GET-DMA_CustomerStatistics.ps1 -Ticket $BTTicket -CustomerId '12345678-0000-0000-0000-000000000000' -Csv '.\output.csv' -PageSize 100
    Runs the script and outputs the DMA statistics for customer with id 12345678-0000-0000-0000-000000000000.
#>

param(
    # Ticket for authentication
    [Parameter(Mandatory=$True)]
    [ManagementProxy.ManagementService.Ticket] $Ticket,
   
    # The id of the customer to work with
    [Parameter(Mandatory=$True)]
    [guid] $CustomerId,

    # The csv output file name
    [Parameter(Mandatory=$False)]
    [string] $Csv = ".\HealthCheckO365-Report-$CustomerId.csv",

    # The batch size of every retrieve request 
    [Parameter(Mandatory=$False)]
    [int] $PageSize = 100,
    
    # The environment to work with
    [Parameter(Mandatory=$False)]
    [ValidateSet("BT", "China")]
    [string] $Env = "BT"
) 

# Retrieve the customer
$customer = Get-BT_Customer -Ticket $Ticket -Environment $Env -FilterBy_Guid_Id $CustomerId
$organizationId = $customer.OrganizationId
if (-not $organizationId) {
    Write-Error "Customer $CustomerId does not exist for given ticket, aborted."
    return
}

# Bind the organization Id to the ticket
if (-not $Ticket.IsPrivileged) 
{
    $Ticket = Get-BT_Ticket -Ticket $Ticket -OrganizationId $organizationId -Environment $Env
    Write-Verbose "Currently non-privilege is used, the page size is limited to 100."
    $PageSize = 100
}

# Log Checkpoint
Write-Host "Health check O365 report will be generated to $Csv for customer $($customer.CompanyName)."

# Retrieve the customer end users with pagination
Write-Host "Retrieving all customerEndUsers...This may take a while."
$count = 0
$customerEndUsers = New-Object System.Collections.ArrayList
While($true)
{    
    [array]$temp = Get-BT_CustomerEndUser -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Boolean_IsDeleted $False -PageOffset $($count*$PageSize) -PageSize $PageSize
    $customerEndUsers.AddRange($temp)
    if ($customerEndUsers.Count % 1000 -eq 0) { Write-Verbose "$($customerEndUsers.Count) entities retrieved." }
    if ($temp.count -lt $PageSize) { break } 
    $count++
}
Write-Host "Total of $($customerEndUsers.Count) customerEndUsers retrieved, gathering information."

# Build the output array list
$outputObjects = New-Object System.Collections.ArrayList

# Build a customer device dictionary to save requests
$customerDevicesDictionary = @{}

# Iterate through the customer end users
foreach ($endUser in $customerEndUsers)
{
   # Retrieve the customer device users with pagination
    Write-Verbose "Retrieving all customerDeviceUsers with end user id $($endUser.Id)...This may take a while."
    $count = 0
    $customerDeviceUsers = New-Object System.Collections.ArrayList
    While($true)
    {            
        [array]$temp = Get-BT_CustomerDeviceUser -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Guid_EndUserId $endUser.Id -FilterBy_Boolean_IsDeleted $False -PageOffset $($count*$PageSize) -PageSize $PageSize
        if (!$temp) { break } 
        $customerDeviceUsers.AddRange($temp)
        if ($customerDeviceUsers.Count % 1000 -eq 0) { Write-Verbose "$($customerDeviceUsers.Count) entities retrieved." }
        if ($temp.count -lt $PageSize) { break } 
        $count++
    }
    Write-Verbose "Totally $($customerDeviceUsers.Count) customerDeviceUsers retrieved."

    # Iterate through the customer device users
    foreach ($deviceUser in $customerDeviceUsers)
    {
        # Retrieve the customer devices
        Write-Verbose "Retrieving the customerDevice with device id $($deviceUser.DeviceId)."   

        # First check the local dictionary
        if ($customerDevicesDictionary.ContainsKey($deviceUser.DeviceId))
        {
            $device = $customerDevicesDictionary[$deviceUser.DeviceId]
        }
        # If not, create a request and cache it in the dictionary
        else
        {
            $device = Get-BT_CustomerDevice -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Guid_Id $deviceUser.DeviceId -FilterBy_Boolean_IsDeleted $False
            if (-not $device) {
                Write-Warning "No matching customer device found for customer device user $($deviceUser.DeviceId)."
                continue
            }
            $customerDevicesDictionary.Add($deviceUser.DeviceId, $device)
        }

        # Process information about incompatible machine items by searching for the different values in the incompatible items
        if ($device.OfficeIncompatibilitySummary)
        {
            $incompatibleItems = $device.OfficeIncompatibilitySummary.Items
            $IsOperatingSystemCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’OSName‘})
            $IsTotalMemoryCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’RAMSizeTotal‘})
            $IsFreeDiskSpaceCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’FreeSpace‘})
            $IsDownloadBandwidthCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’DownloadSpeed‘})
            $IsUploadBandwidthCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’UploadSpeed‘})
        }
        else
        {
            $IsOperatingSystemCompatible = $True
            $IsTotalMemoryCompatible = $True
            $IsFreeDiskSpaceCompatible = $True
            $IsDownloadBandwidthCompatible = $True
            $IsUploadBandwidthCompatible = $True
        }

        # Process information about incompatible software items by searching for the different values in the incompatible items
        if ($deviceUser.OfficeIncompatibilitySummary)
        {
            $incompatibleItems = $deviceUser.OfficeIncompatibilitySummary.Items
            $IsInternetExplorerCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’InternetExplorerVersion‘})
            $IsFirefoxCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’FirefoxVersion‘})
            $IsChromeCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’ChromeVersion‘})
            $IsEdgeCompatible = -not ($incompatibleItems | ? {$_.Type -eq ’EdgeVersion‘})
        }
        else
        {
            $IsInternetExplorerCompatible = $True
            $IsFirefoxCompatible = $True
            $IsChromeCompatible = $True
            $IsEdgeCompatible = $True
        }

        # Retrieve customer device user
        Write-Verbose "Retrieving customerDeviceUsers with end user id $($endUser.Id) and device id $($deviceUser.DeviceId)."
        $deviceUser = Get-Bt_CustomerDeviceUser -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Guid_EndUserId $endUser.Id -FilterBy_Guid_DeviceId $deviceUser.DeviceId -FilterBy_Boolean_IsDeleted $False

        # Start processing data
        Write-Verbose "Generating a new item in the report based on customerDevice and deviceUser."

        # Create a new item for the summary
        $object = New-Object -TypeName PSObject
        $object | Add-Member -MemberType NoteProperty -Name CustomerName -Value $customer.CompanyName
        $object | Add-Member -MemberType NoteProperty -Name UserPrimaryEmailAddress -Value $endUser.PrimaryEmailAddress

        # Add information about the device-level fields in the summary
        $object | Add-Member -MemberType NoteProperty -Name DeviceName -Value $device.DeviceName
        $object | Add-Member -MemberType NoteProperty -Name OperatingSystemName -Value $device.OSName
        $object | Add-Member -MemberType NoteProperty -Name IsOperatingSystemCompatible -Value $IsOperatingSystemCompatible
        $object | Add-Member -MemberType NoteProperty -Name Processor -Value $device.ProcessorName      
        $object | Add-Member -MemberType NoteProperty -Name TotalMemory -Value $device.RAMSizeTotal
        $object | Add-Member -MemberType NoteProperty -Name IsTotalMemoryCompatible -Value $IsTotalMemoryCompatible         
        $object | Add-Member -MemberType NoteProperty -Name TotalDiskSpace -Value $device.DiskSpaceTotal               
        $object | Add-Member -MemberType NoteProperty -Name FreeDiskSpace -Value $device.DiskSpaceFree
        $object | Add-Member -MemberType NoteProperty -Name IsFreeDiskSpaceCompatible -Value $IsFreeDiskSpaceCompatible

        # Add information about the device user-level fields in the summary
        if ($deviceUser)
        {
            $object | Add-Member -MemberType NoteProperty -Name AgentStatus -Value $deviceUser.AgentStatus
            $object | Add-Member -MemberType NoteProperty -Name AgentVersion -Value $deviceUser.AgentVersion
            $object | Add-Member -MemberType NoteProperty -Name AgentLastHeartbeat -Value $deviceUser.LastAgentHeartbeatDate
        }

        # Add information about the bandwidth
        $object | Add-Member -MemberType NoteProperty -Name DownloadBandwidth -Value $device.DownloadSpeed       
        $object | Add-Member -MemberType NoteProperty -Name IsDownloadBandwidthCompatible -Value $IsDownloadBandwidthCompatible
        $object | Add-Member -MemberType NoteProperty -Name UploadBandwidth -Value $device.UploadSpeed
        $object | Add-Member -MemberType NoteProperty -Name IsUploadBandwidthCompatible -Value $IsUploadBandwidthCompatible      

        # Add information about the device user-level fields in the summary
        $object | Add-Member -MemberType NoteProperty -Name IsOfficeVersionCompatible -Value $deviceUser.IsOfficeCompatible
        $object | Add-Member -MemberType NoteProperty -Name InternetExplorerVersion -Value $deviceUser.InternetExplorerVersion
        $object | Add-Member -MemberType NoteProperty -Name IsInternetExplorerCompatible -Value $IsInternetExplorerCompatible        
        $object | Add-Member -MemberType NoteProperty -Name FirefoxVersion -Value $deviceUser.FirefoxVersion
        $object | Add-Member -MemberType NoteProperty -Name IsFirefoxCompatible -Value $IsFirefoxCompatible
        $object | Add-Member -MemberType NoteProperty -Name ChromeVersion -Value $deviceUser.ChromeVersion
        $object | Add-Member -MemberType NoteProperty -Name IsChromeCompatible -Value $IsChromeCompatible
        $object | Add-Member -MemberType NoteProperty -Name EdgeVersion -Value $deviceUser.EdgeVersion
        $object | Add-Member -MemberType NoteProperty -Name IsEdgeCompatible -Value $IsEdgeCompatible
        $object | Add-Member -MemberType NoteProperty -Name IsO365CompatibleClients -Value $deviceUser.IsOfficeReadyForClients
        $object | Add-Member -MemberType NoteProperty -Name IsO365CompatibleWebApps -Value $deviceUser.IsOfficeReadyForWebApps
                
        # Append to output
        [void]$outputObjects.Add($object)
    }
}

# Log Checkpoint
Write-Verbose "Outputing csv..."

# Iterate through the $outputObjects and write report
$outputObjects | Export-Csv $Csv
