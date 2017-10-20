param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath)

# Initialize the context
.\Init.ps1

# Import the Runbook from FilePath
Import-BT_OfferingMetadata -Ticket $mspc.WorkgroupTicket -FilePath $FilePath