<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			AssignSubscription.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.00
	Date:			April 26, 2017

	Windows Version:	WINDOWS 10 ENTERPRISE

	Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

	Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
	
.SYNOPSIS
	Submit apply subscription request for all the users within the specified workgroup.

.DESCRIPTION 	
	This script is an example to show the subscription assignment is programmable. This script will create a add subscription task for each customer end user within the workgroup.
    The task will be handled by AEF.

.INPUTS
	-[ManagementProxy.ManagementService.Ticket] Ticket, the ticket for authentication.
    -[guid] WorkGroupId, the id of the workgroup.
	-[int] PageSize, the batch size of every retrieve request.
    -[string] Env, the context to work with. Valid options : BT, China.

.EXAMPLE
  	.\AssignSubscription.ps1 -Ticket -WorkGroupId 
	Runs the script to submit add subscription tasks for each one of the customer end users within the workgroup.
#>

param(
    # Ticket 
	[Parameter(Mandatory=$True)]
	[ManagementProxy.ManagementService.Ticket] $Ticket,

    # Ticket for authentication
    [Parameter(Mandatory=$True)]
    [guid] $WorkGroupId,

    # The batch size of every retrieve request 
    [Parameter(Mandatory=$False)]
    [int] $PageSize = 100,

    # The environment to work with
    [Parameter(Mandatory=$False)]
    [ValidateSet("BT", "China")]
    [string] $Env = "BT"
) 

# Check if the ticket is privileged, if not, cap the page size
if (-not $Ticket.IsPrivileged) 
{
    Write-Verbose "Currently non-privilege is used, the page size is limited to 100."
    $PageSize = 100
}

# Get workgroup
$workgroup = Get-BT_Workgroup -Ticket $ticket -FilterBy_Guid_Id $WorkGroupId -Environment $Env -FilterBy_Boolean_IsDeleted $False
if (-not $workgroup) {
    Write-Error "Workgroup $WorkGroupId does not exist for given ticket, aborted."
    return
}

# Get a list of customers under the workgroup with pagination
$count = 0
$customers = New-Object System.Collections.ArrayList
While($true)
{    
    [array]$temp =  Get-BT_Customer -Ticket $ticket -FilterBy_Guid_WorkgroupId $workgroup.Id -Environment $Env -FilterBy_Boolean_IsDeleted $False -PageOffset $($count*$PageSize) -PageSize $PageSize
    $customers.AddRange($temp)
    if ($temp.count -lt $PageSize) { break } 
    $count++
}

# Get another ticket with both organization Id and workgroup Id
$upgradedTicket = Get-BT_Ticket -Ticket $ticket -WorkgroupId $workgroup.Id -OrganizationId $workgroup.WorkgroupOrganizationId -Environment $Env -KeepPrivileged

# Process each customer
foreach($customer in $customers)
{
    # Get a list of customer end users with pagination
    $count = 0
    $customerEndUsers = New-Object System.Collections.ArrayList
    While($true)
    {    
        [array]$temp =  Get-BT_CustomerEndUser -Ticket $ticket -FilterBy_Guid_OrganizationId $customer.OrganizationId -Environment $Env -FilterBy_Boolean_IsDeleted $False -PageOffset $($count*$PageSize) -PageSize $PageSize
        $customerEndUsers.AddRange($temp)
        if ($temp.count -lt $PageSize) { break } 
        $count++
    }

    # Get the product sku id
    $productId = Get-BT_ProductSkuId -Ticket $ticket -ProductName MspcEndUserYearlySubscription -Environment $Env 

    # Assign subscription to each customer end user
    # ProductSkuId used here is the id of 1-year subscription product
    $customerEndUsers | %{ Add-BT_Subscription -Ticket $upgradedTicket -SubscriptionEntityReferenceType CustomerEndUser -EntityReferenceId $_.Id -ProductSkuId $productId -Environment $Env -WorkgroupOrganizationId $workgroup.WorkgroupOrganizationId }
}