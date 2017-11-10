# Initialize the context
.\Init.ps1

# Here are 3 common ways to retrieve entities
# 1. Retrieve all the customers in a workgroup and process them one by one
# Use -RetrieveAll and piping with ForEach 
Get-BT_Customer -Ticket $mspc.WorkgroupTicket -RetrieveAll | ForEach {
    # Process each customer retrieved
    Write-Host $_.CompanyName
}

# 2. Retrieve and process customers from a workgroup in batches
# Use paging
$pageSize = 100
$count = 0
While( $true ) {   
    # Retrieve customers in batch
    [array]$batch = Get-BT_Customer -Ticket $mspc.WorkgroupTicket -PageOffset $($count * $pageSize)
    if ( -not $batch ) { 
        break
    }
    
    # Update each customer in the batch
    ForEach($customer in $batch) {
        $customer.CountryName = 'USA'
    }

    # Send update request
    Set-BT_Customer -Ticket $mspc.WorkgroupTicket -Customer $batch

    # Increase count
    $count ++
}

# 3. Retrieve all the customers in a workgroup to get certain info, e.g. count 
# Use -RetrieveAll
# Note: This is the least efficient way among the 3 cases since it loads all the entities into the memory
$customers = Get-BT_Customer -Ticket $mspc.WorkgroupTicket -RetrieveAll
$customers.Count