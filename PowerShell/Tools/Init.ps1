# Initialization script that emulates environment for automated tasks
param (
    [switch] $Clear,
    [switch] $Customer
)

# If $Clear option is specified, clear saved state and exit
if ($Clear) {
    $global:mspc = $null
    $global:creds = $null
    Write-Host "Saved MSPComplete context has been cleared."
    return
}

# Enable logging
$InformationPreference = 'Continue'
$DebugPreference = 'Continue'

# Logging
Write-Information "Initializing BitTitan Automated Task Environment"
Write-Debug "Customer: $Customer."

# Load BitTitan PowerShell module, if it has not been loaded already
if (-not (Get-Module "BitTitanPowerShell")) {
    Import-Module 'C:\Program Files (x86)\BitTitan\BitTitan PowerShell\BitTitanPowerShell.dll' -ErrorAction Stop
}

# Initialize the context variables
if (-not $global:mspc) {
    # Initialize the context
    $global:mspc = @{}
    $mspc.Data = @{}
    
    # Get credentials
    if (-not $global:creds) {
        $global:creds = Get-Credential -Message "Enter BitTitan credentials"
    }
   
    # Initialize the base ticket
    $mspc.Ticket = Get-BT_Ticket -Credentials $creds -ServiceType BitTitan
     
    # Prompt for workgroup
    $workgroupId = [GUID](Read-Host -Prompt 'Workgroup ID')

    # Retrieve customer and workgroup
    if ($Customer) {
        $customerId = [GUID](Read-Host -Prompt 'Customer ID')    
        $mspc.Customer = Get-BT_Customer -Ticket $mspc.Ticket -Id $customerId
    }
    $mspc.Workgroup = Get-BT_Workgroup -Ticket $mspc.Ticket -Id $workgroupId

    # Initialize the additional tickets
    if ($Customer) {
        $mspc.CustomerTicket = Get-BT_Ticket -Ticket $mspc.Ticket -OrganizationId $mspc.Customer.OrganizationId
    }
    $mspc.WorkgroupTicket = Get-BT_Ticket -Ticket $mspc.Ticket -OrganizationId $mspc.Workgroup.WorkgroupOrganizationId
    $mspc.MigrationWizTicket = Get-MW_Ticket -Credentials $creds
}
