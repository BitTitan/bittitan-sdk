# Initialize the context
# Create a customer scoped ticket
.\Init.ps1 -Customer

# Here are 3 common ways to retrieve entities
# 1. Retrieve all the endpoints under a customer and process them one by one
# Use -RetrieveAll and piping with ForEach 
Get-BT_Endpoint -Ticket $mspc.CustomerTicket -RetrieveAll | ForEach {
    # Process each endpoint retrieved
    Write-Host $_.Name
}

# 2. Retrieve and process endpoints under a customer in batches
# Use paging
$pageSize = 100
$count = 0
While( $true ) {   
    # Retrieve endpoints in batch
    [array]$batch = Get-BT_Endpoint -Ticket $mspc.CustomerTicket -PageOffset $($count * $pageSize)
    if ( -not $batch ) { 
        break
    }
    
    # Update each endpoint in the batch
    ForEach($endpoint in $batch) {
        $endpoint.Name += "_test"
    }

    # Send update request
    Set-BT_Endpoint -Ticket $mspc.CustomerTicket -Endpoint $batch

    # Increase count
    $count ++
}

# 3. Retrieve all the endpoints under a customer to get certain info, e.g. count 
# Use -RetrieveAll
# Note: This is the least efficient way among the 3 cases since it loads all the entities into the memory
$endpoints = Get-BT_Endpoint -Ticket $mspc.CustomerTicket -RetrieveAll
$endpoints.Count