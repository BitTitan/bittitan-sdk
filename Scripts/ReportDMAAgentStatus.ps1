<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			ReportDMAAgentStatus.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.00
	Date:			March 30, 2017

	Exchange Version:	2016
	Windows Version:	WINDOWS 10 ENTERPRISE

	Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

	Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
	
.SYNOPSIS
	Reports the DMA agent status per device, possible status could be installing, uninstalling, uninstalled, running, etc.

.DESCRIPTION 	
	Retrieves the customerDevices for given customer and report the agent status.

.INPUTS
	-[ManagementProxy.ManagementService.Ticket] Ticket, the ticket for authentication.
	-[guid] CustomerId, the id of the customer to report.
	-[string] Csv, the csv output path.
	-[string] Env, the context to work with. Valid options : BT, China.

.EXAMPLE
  	.\ReportDMAAgentStatus.ps1 -Ticket -customerId '12345678-0000-0000-0000-000000000000'
#>
param(    
    # Ticket 
	[Parameter(Mandatory=$True)]
	[ManagementProxy.ManagementService.Ticket] $Ticket,

    # Customer Id
	[Parameter(Mandatory=$True, ValueFromPipeline=$True)]
	[guid] $CustomerId,

	# Csv
	[Parameter(Mandatory=$False)]
	[string] $Csv = ".\DMAStatusReport-$CustomerId.csv",

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
Write-Verbose "Retrieving customerDevices...This may take a while..."
$count = 0
$customerDevices = New-Object System.Collections.ArrayList
While($true)
{   
    # Retrieve a batch of customer device entities
    [array]$temp = Get-BT_CustomerDevice -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Boolean_IsDeleted $False -PageOffset $($count*$pageSize) -PageSize $pageSize

    # Break the loop if nothing is retrieved
    if (-not $temp) { break }

    # Add the customer devices
    $customerDevices.AddRange($temp)

    # Break the loop if it is the last batch
    if ($temp.count -lt $pageSize) { break }
    
    # Batch count increase 
    $count++   
}
Write-Verbose "Totally $($Count*$pageSize + $temp.Count) customerDevices are retrieved."

# Log Checkpoint
Write-Verbose "Outputing csv..."

# Iterate through the $outputObjects and write report
$customerDevices | Select DeviceName, Id, AgentStatus | Export-Csv $Csv
