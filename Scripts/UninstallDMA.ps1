<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			UninstallDMA.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.00
	Date:			March 28, 2017

	Windows Version:	WINDOWS 10 ENTERPRISE

	Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

	Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
	
.SYNOPSIS
	Updates the desired state of all customerDevices to uninstalled.

.DESCRIPTION 	
	To uninstall DMA, we need to update the desired state of corresponding customerDevice entity. This script takes in a customer id and flips all the customerDevice entities 
    for the given customer. The environment is restricted to production. 

.INPUTS
	-[ManagementProxy.ManagementService.Ticket] Ticket, the ticket for authentication.
	-[guid] CustomerId, the id of the customer to report.
	-[string] Env, the context to work with. Valid options : BT, China.

.EXAMPLE
  	.\UninstallDMA.ps1 -Ticket -customerId '12345678-0000-0000-0000-000000000000'
#>
param(    
    # Ticket 
    [Parameter(Mandatory=$True)]
    [ManagementProxy.ManagementService.Ticket] $Ticket,

    # Customer Id
    [Parameter(Mandatory=$True, ValueFromPipeline=$True)]
    [guid] $CustomerId,
	
	# Env
	[Parameter(Mandatory=$False)]
    [ValidateSet("BT", "China")]
    [string] $Env = "BT"
) 

# Retrieve the customer
$customer = Get-BT_Customer -Ticket $Ticket -Environment $Env -FilterBy_Guid_Id $CustomerId
$organizationId = $customer.OrganizationId
if (-not $organizationId) {
    Write-Error "Customer does not exist for the given customer Id $CustomerId, aborted."
    return
}
Write-Verbose "Working with customer $($customer.CompanyName)"

# Bind the organization Id to the ticket
if (-not $Ticket.IsPrivileged) 
{
    $Ticket = Get-BT_Ticket -Ticket $Ticket -OrganizationId $organizationId -Environment $Env
}

# Set page size
$pageSize = 100

# Retrieve and update the customer devices with pagination
Write-Verbose "Updating all customerDevices...This may take a while..."
$count = 0
While($true)
{   
    # Retrieve a batch of customer device entities
    [array]$temp = Get-BT_CustomerDevice -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Boolean_IsDeleted $False -FilterBy_DesiredState Installed -PageOffset $($count*$pageSize) -PageSize $pageSize

    # Break the loop if nothing is retrieved
    if (-not $temp) { break }

    # Update the desired state to uninstall for all the customer device entities
    [void](Set-BT_CustomerDevice -Ticket $Ticket -Environment $Env -customerdeviceArray $temp -DesiredState Uninstalled)

    # Break the loop if it is the last batch
    if ($temp.count -lt $pageSize) { break }
    
    # Batch count increase 
    $count++   
}
Write-Verbose "Totally $($Count*$pageSize + $temp.Count) customerDevices are retrieved and planned to uninstall."