<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			CheckTaskInstanceStatus.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.00
	Date:			DECEMBER 8, 2016

	Exchange Version:	2016
	Windows Version:	WINDOWS 10 ENTERPRISE

	Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

	Copyright: 		Copyright © 2016 BitTitan. All rights reserved.
	
.SYNOPSIS
	Gets the status for given Task instance.

.DESCRIPTION 	
	This script tries to retrieve the task instance by input id, and returns its status. 

.INPUTS
    -[ManagementProxy.ManagementService.Ticket] Ticket, the ticket for authentication.
    -[guid] TaskInstanceId, the id of the task instance to query.
    -[string] Env, the context to work with. Valid options : BT, China.

.OUTPUTS
	[ManagementProxy.ManagementService.TaskStatus]
	
.EXAMPLE
  	.\CreatePublicFolderPermissions.ps1 -Ticket -TaskInstanceId 
	Runs the script and outputs the status of the specified task instance.
#>

param(
    [Parameter(Mandatory=$True)]
    [ManagementProxy.ManagementService.Ticket] $Ticket,
    [Parameter(Mandatory=$True)]
    [guid] $TaskInstanceId,
    [Parameter(Mandatory=$False)]
    [ValidateSet("BT", "China")]
    [string] $Env = "BT"
) 

# Get the task instance
$task = Get-BT_TaskInstance -Ticket $Ticket -Environment $Env -FilterBy_Guid_Id $TaskInstanceId -FilterBy_Boolean_IsDeleted $false
if (-NOT $task)
{
    Write-Error "Task instance with id {$TaskInstancId} does not exist."
    return
}

# Check the status of the task
Write-Output $task.Status