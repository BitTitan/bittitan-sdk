<#
.NOTES
    Company:		BitTitan, Inc.
    Title:			CreateAndRunRemoveCalendarItemsForOffboardingO365User.PS1
    Author:			SUPPORT@BITTITAN.COM
    Requirements: 
    
    Version:		1.00
    Date:			March 31, 2017

    Windows Version:	WINDOWS 10 ENTERPRISE

    Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

    Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
    
.SYNOPSIS
    Instantiates and executes a Remove Calendar Items For Offboarding O365 User service.

.DESCRIPTION 	
    This script is an example to show the creation and execution of MSPC services is programmable. This script will create a Remove Calendar Items For Offboarding O365 User offering instance with related task instances and tasktargets, complete the first two 
    tasks with the input parameters, and submit the last task to AEF.

.INPUTS
    -[ManagementProxy.ManagementService.Ticket] Ticket, the ticket for authentication.
    -[guid] CustomerId, the id of the customer.
    -[string] EndPointName, the name of the endpoint to work with.
    -[string] EndUserEmailAddressToOffboard, the end user to off board.
    -[string] EndUserEmailAddressToRelocate, the end user to assign offboarded user resources to.
    -[int] PageSize, the batch size of every retrieve request.
    -[string] Env, the context to work with. Valid options : BT, China.

.EXAMPLE
    .\CreatePublicFolderPermissions.ps1 -Ticket -EndPointName -UserIdList -Env -CustomerId
    Runs the script to instantiate and execute an O365 password reset service for the specific customer.
#>

param(
    # Ticket for authentication
    [Parameter(Mandatory=$True)]
    [ManagementProxy.ManagementService.Ticket] $Ticket,

    # The Id of the cusomter
    [Parameter(Mandatory=$True)]
    [guid] $CustomerId,

    # The name of the endpoint to use
    [Parameter(Mandatory=$True)]
    [string] $EndPointName,

    # The email address of user to offboard
    [Parameter(Mandatory=$True)]
    [string] $EndUserEmailAddressToOffboard,

    # The email address of user to assign offboarded user resources to
    [Parameter(Mandatory=$True)]
    [string] $EndUserEmailAddressToRelocate,

    # The batch size of every retrieve request 
    [Parameter(Mandatory=$False)]
    [int] $PageSize = 100,

    # The environment to work with
    [Parameter(Mandatory=$False)]
    [ValidateSet("BT", "China")]
    [string] $Env = "BT"
) 

# Validate Ticket
if ([GUID]::Empty -ne $Ticket.OrganizationId)
{
    Write-Error "An unscoped ticket is needed."
    return    
}

# Validate customer scope
$customer = Get-BT_Customer -Ticket $Ticket -Environment $Env -FilterBy_Guid_Id $CustomerId -FilterBy_Boolean_IsDeleted $False
if (-not $customer)
{
    Write-Error "There does not exist the customer with id $CustomerId."
    return
}

# Bind the ticket to the customer scope
$scopedTicket = Get-BT_Ticket -Ticket $Ticket -Environment $Env -OrganizationId $customer.OrganizationId
if (-not $scopedTicket)
{
    Write-Error "Failed to scope the ticket with customer organization."
    return
}

# Validate input endpoint exists
$endPoint = Get-BT_Endpoint -Ticket $scopedTicket -Environment $Env -FilterBy_String_Name $EndPointName -FilterBy_Boolean_IsDeleted $False
if (-NOT $endPoint) 
{ 
    Write-Error "There does not exist an endpoint with name $EndPointName."
    return 
}

# Retrieve the user to offboard
$userToOffboard = Get-BT_CustomerEndUser -Ticket $scopedTicket -Environment $Env -FilterBy_String_PrimaryEmailAddress $EndUserEmailAddressToOffboard
if (-not $userToOffboard) 
{ 
    Write-Error "There does not exist an user with email address $EndUserEmailAddressToOffboard."
    return
}

# Retrieve the user assign resources
$userToRelocate = Get-BT_CustomerEndUser -Ticket $scopedTicket -Environment $Env -FilterBy_String_PrimaryEmailAddress $EndUserEmailAddressToRelocate
if (-not $userToRelocate) 
{ 
    Write-Error "There does not exist an user with email address $EndUserEmailAddressToRelocate."
    return
}

# Retrieve offering metadata id for O365 User Password Reset
$offeringMetadata = Get-BT_OfferingMetadata -Ticket $Ticket -Environment $Env -FilterBy_String_KeyName "Office365RemoveCalendarItems"
$offeringMetadataId = $offeringMetadata.Id

# Create a new Offering instance and deploy tasks
[void]($offeringInstance = Add-BT_OfferingInstance -Ticket $scopedTicket -Environment $Env -OfferingMetadataId $offeringMetadataId)
if (-not $offeringInstance)
{
    Write-Error "Fail to create offering."
    return
}

# Retrieve the task instances with pagination
$count = 0
$taskInstances = New-Object System.Collections.ArrayList
While ($true)
{    
    [array]$temp = Get-BT_TaskInstance -Ticket $scopedTicket -Environment $Env -FilterBy_Guid_OfferingInstanceId $offeringInstance.Id -PageOffset $($count*$PageSize) -PageSize $PageSize -FilterBy_Boolean_IsDeleted $False
    if ($temp) 
    { 
        $taskInstances.AddRange($temp)
        if ($temp.count -lt $PageSize) 
        { 
            break 
        }
    }
    else { break }
    $count++
}
Write-Verbose "Totally $($taskInstances.Count) tasks retrieved."

# Order the task instances by Execution order         
if (-not $taskInstances)
{
    Write-Error "Fail to create tasks."
    return
}                                                                                                                                                                                                             
$taskInstances = $taskInstances | Sort-Object ExecutionOrder

################### The 1st (out of 4) task: Select O365 endpoint #####################
# Create the taskTarget
[void](Add-BT_TaskTarget -Ticket $scopedTicket -Environment $Env -TaskInstanceId $taskInstances[0].Id -ReferenceEntityType 'EndPoint' -ReferenceEntityId $endPoint.Id)

# Perform the complete action
Complete-BT_TaskInstance -Ticket $scopedTicket -Environment $Env -TaskInstanceId $taskInstances[0].Id

####################### The 2nd (out of 4) task: Select a user to offboard #########################
# Create the taskTargets
[void](Add-BT_TaskTarget -Ticket $scopedTicket -Environment $Env -TaskInstanceId $taskInstances[1].Id -ReferenceEntityType 'CustomerEndUser' -ReferenceEntityId $userToOffboard.Id)

# Perform the complete action
Complete-BT_TaskInstance -Ticket $scopedTicket -Environment $Env -TaskInstanceId $taskInstances[1].Id

################### The 3rd (out of 4) task: Select a user to assign offboarded user resources to #####################
# Create the taskTargets
[void](Add-BT_TaskTarget -Ticket $scopedTicket -Environment $Env -TaskInstanceId $taskInstances[2].Id -ReferenceEntityType 'CustomerEndUser' -ReferenceEntityId $userToRelocate.Id)

# Perform the complete action
Complete-BT_TaskInstance -Ticket $scopedTicket -Environment $Env -TaskInstanceId $taskInstances[2].Id

############# The 4th (out of 4） task: Remove Calendar Items for offboarding an Office 365 userComplete #################
# Perform the complete action for AEF task
Complete-BT_TaskInstance -Ticket $scopedTicket -Environment $Env -TaskInstanceId $taskInstances[3].Id

################## End of Service: Remove Calendar Items For Offboarding O365 User ########################