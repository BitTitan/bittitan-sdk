<#
.NOTES
    Company:          BitTitan, Inc.
    Title:            Office 365 Groups Bulk Endpoint Creator.ps1
    Author:           Support@BitTitan.com

    Version:          1.00
    Date:             August 17, 2018

    Disclaimer:       This script is provided 'AS IS'. No warranty is provided either expressed or implied

    Copyright:        Copyright © 2018 BitTitan. All rights reserved

.SYNOPSIS
    Bulk generates Office 365 endpoints, projects, and items, based on input from a CSV file.

.DESCRIPTION
    This script will read in a CSV file containing endpoint names, Group URLs, admin usernames and admin passwords, and will
    automatically generate endpoints as well as projects and project items. It will associate corresponding source and destination
    endpoints with each project when it creates the project.
    Fields that need to be included in the CSV (in this order): Source Endpoint Name, Source Group URL, Source Admin Username, Source Admin Password, Destination Endpoint Name, Destination Group URL, Destination Admin Username, Destination Admin Password


#>

# Reads a CSV file.
function Read-Csv($csvPath)
{
    $csvFile = Import-Csv -Path $csvPath
    return $csvFile
}

# Gets the workgroup that the data will be stored under.
function Get-Workgroup($ticket)
{
    $workgroups = Get-BT_Workgroup -Ticket $ticket
    
    if ($workgroups.Length -gt 1)
    {
        # Prompt which workgroup to to use if there are more than 1.
        Write-Host "You have multiple workgroups. Please make a selection:"

        for ($i = 0; $i -lt $workgroups.Length; $i++)
        {
            Write-Host "$i) Id:"$workgroups[$i].Id"- Name:"$workgroups[$i].Name
        }

		$workgroupChoice = Read-Host
		return $workgroups[$workgroupChoice]
    }
    elseif ($workgroups.Length -eq 1)
    {
        return $workgroups[0]
    }
    else
    {
        return $false
    }
}

# Gets the customer that will contain the endpoints and projects.
function Get-Customer($ticket, $workgroupIdNum)
{
	$customers = Get-BT_Customer -Ticket $ticket -WorkgroupId $workgroupIdNum
	
    if ($customers.Length -gt 1)
	{
        # Prompt which customer to use if there are more than 1.
		Write-Host "You have multiple customers. Please make a selection:"

		for ($i = 0; $i -lt $customers.Length; $i++)
		{
			Write-Host "$i) Id:"$customers[$i].OrganizationId"- Name:"$customers[$i].CompanyName
		}

		$customersChoice = Read-Host
		return $customers[$customersChoice]
	}
	elseif ($customers.Length -eq 1)
	{
		return $customers[0]
	}
	else
	{
		return $false
	}
}

# Generates the Office 365 Group endpoints.
function Add-Endpoints($endpoints, $ticket, $customerOrgId, $sourceOrDestination)
{
    $allEndpoints = @()

    $scopedTicket = Get-BT_Ticket -Ticket $ticket -OrganizationId $customerOrgId

    for ($i = 0; $i -lt ($endpoints.Length); $i++)
    {
        $endpointConfiguration = New-Object ManagementProxy.ManagementService.SharePointConfiguration -Property @{
            "AdministrativeUsername" = $endpoints[$i].("$sourceOrDestination Admin Username");
            "AdministrativePassword" = $endpoints[$i].("$sourceOrDestination Admin Password");
            "UseAdministrativeCredentials" = $true;
            "Url" = $endpoints[$i].("$sourceOrDestination Group URL");
        }

        $endpointCreated = Add-BT_Endpoint -Ticket $scopedTicket -Configuration $endpointConfiguration -Type Office365Groups -Name $endpoints[$i].("$sourceOrDestination Endpoint Name")
        Write-Host "Endpoint"$endpointCreated.Name "was created successfully."
        
        $allEndpoints += $endpointCreated
    }

    return $allEndpoints
}

# Creates the migration projects, using the endpoint information passed in.
function Add-Projects($endpoints, $sourceEndpoints, $destinationEndpoints, $ticket, $customerOrgId)
{
    $projectIds = @()

    for ($i = 0; $i -lt $endpoints.Length; $i++)
    {
        $sourceConfig = New-MW_SharePointConfiguration -UseAdministrativeCredentials $true -AdministrativeUsername $endpoints[$i].'Source Admin Username' -AdministrativePassword $endpoints[$i].'Source Admin Password' -Url $endpoints[$i].'Source Group Url'
        $destinationConfig = New-MW_SharePointConfiguration -UseAdministrativeCredentials $true -AdministrativeUsername $endpoints[$i].'Destination Admin Username' -AdministrativePassword $endpoints[$i].'Destination Admin Password' -Url $endpoints[$i].'Destination Group Url'

        $currentProject = Add-MW_MailboxConnector -Ticket $ticket -ProjectType Storage -ExportType Office365Groups -ImportType Office365Groups -Name $sourceEndpoints[$i].Name -UserId $ticket.UserId -OrganizationId $customerOrgId -SelectedExportEndpointId $sourceEndpoints[$i].Id -SelectedImportEndpointId $destinationEndpoints[$i].Id -ExportConfiguration $sourceConfig -ImportConfiguration $destinationConfig -MaximumItemFailures 100 -AdvancedOptions 'FolderLimit=20000 InitializationTimeout=28800000'
        $projectIds += $currentProject.Id

        Write-Host 'Project' $sourceEndpoints[$i].Name 'was created successfully.'
    }

    return $projectIds
}

# Adds the default line item to the project.
function Add-ItemsToProject($ticket, $projectIds)
{
    $projectItems = @()

    for ($i = 0; $i -lt $projectIds.Length; $i++)
    {
        $projectItems += Add-MW_Mailbox -Ticket $ticket -ConnectorId $projectIds[$i] -ExportLibrary "Shared Documents" -ImportLibrary "Shared Documents"
    }

    return $projectItems
}

function Main
{
    # Get the name of the CSV that contains the list of Office 365 Groups
    Write-Host "Please enter the path and name of the CSV containing the endpoints to create"
    $csvPath = Read-Host

    # Read the CSV File
    $csvFile = Read-Csv $csvPath

    # Gather BitTitan credentials
    $creds = Get-Credential -Message "Enter your BitTitan Credentials"
    try
    {
        $btTicket = Get-BT_Ticket -Credentials $creds
    }
    catch
    {
        Write-Host "Could not create the BT Ticket. Press ENTER to exit."
        Read-Host
        return
    }
    try
    {
        $mwTicket = Get-MW_Ticket -Credentials $creds
    }
    catch
    {
        Write-Host "Could not create the MW Ticket. Press ENTER to exit."
        Read-Host
        return
    }
    # Grab the workgroup. If there is more than one workgroup, it will prompt the user.
    $workgroup = Get-Workgroup $btTicket

    # Get the Customer. If there is more than one customer, it will prompt the user.
    $customer = Get-Customer $btTicket $workgroup.Id

    # Create the endpoints based on the information in the CSV File.
    $sourceEndpointsCreated = Add-Endpoints $csvFile $btTicket $customer.OrganizationId "Source"
    $destinationEndpointsCreated = Add-Endpoints $csvFile $btTicket $customer.OrganizationId "Destination"

    # Create the projects and users in MigrationWiz.
    $projectsCreated = Add-Projects $csvFile $sourceEndpointsCreated $destinationEndpointsCreated $mwTicket $customer.OrganizationId

    # Create a default line item within each project.
    $projectItemsCreated = Add-ItemsToProject $mwTicket $projectsCreated
}

Main