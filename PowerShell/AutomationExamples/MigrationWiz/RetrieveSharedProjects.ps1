# Authenticate
$credential = Get-Credential
$ticket = Get-BT_Ticket -ServiceType BitTitan -Credentials $credential

# Retrieve a workgroup and the customers in the workgroup
$workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')
$workgroup = Get-BT_Workgroup -Ticket $ticket -Id $workgroupId
$customers = Get-BT_Customer -Ticket $ticket -WorkgroupId $workgroup.Id

# Get a MW ticket with the project sharing option, for the workgroup retrieved previously
$mwTicketWithProjectSharing = Get-MW_Ticket -Credentials $credential -WorkgroupId $workgroup.Id -IncludeSharedProjects

# Retrieve projects
$projects = Get-MW_MailboxConnector -Ticket $mwTicketWithProjectSharing -OrganizationId $customers.OrganizationId

# Retrieve items under each project
$projects | ForEach {
    Write-Host("`n Project: `"$($_.Name)`"")
    
    # Retrieve project items
    $projectItems = Get-MW_Mailbox -Ticket $mwTicketWithProjectSharing -ConnectorId $_.Id
    if ( -not $projectItems ) { 
        Write-Host("`t0 items")
    }
    else {
        # Retrieve the last migration submitted for each item
        $projectItems | ForEach {
            $projectItemMigration = Get-MW_MailboxMigration -Ticket $mwTicketWithProjectSharing -MailboxId $_.Id -SortBy_CreateDate_Descending -PageSize 1
            
            # Print result
            if ( -not $projectItemMigration ) { 
                Write-Host("`t $($_.ExportEmailAddress): No migrations")
            }
            else {
                Write-Host("`t $($_.ExportEmailAddress): Last migration: $($projectItemMigration.CreateDate), $($projectItemMigration.Status)")
            }            
        }
    }
}
