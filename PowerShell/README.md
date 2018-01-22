# BitTitan PowerShell
This folder contains a set of tools, building blocks and full scenario scripts tailored to work with PowerShell using the BitTitan SDK. 

## Installation
Download the [MSI](https://www.bittitan.com/downloads/bittitanpowershellsetup.msi) from the BitTitan website and run it to install the SDK.

## Quick Startup
1. Install the BitTitan PowerShell SDK. The installation section above details the steps to install the SDK.
2. Clone the [bittitan-sdk repository](https://github.com/BitTitan/bittitan-sdk) or download the desired scripts.
    ```shell
    git clone https://github.com/BitTitan/bittitan-sdk.git
    ```
3. Copy the `Init.ps1` (bittitan-sdk/PowerShell/Tools/Init.ps1) script to the working directory. If the repository was not cloned, download the [Init script](https://github.com/BitTitan/bittitan-sdk/blob/master/PowerShell/Tools/Init.ps1) and copy it to the working directory.
4. Execute desired scripts from the working directory.

## Contents
* `/BuildingBlocks (Building Blocks)` - Samples to help developers build scripts that automate various operations.
* `/AutomationExamples (Automation Examples)` - Example scripts that enable developers to address specific scenarios or self-troubleshoot issues.
* `/AutomatedTaskScripts (Automated Task Scripts)` - Scripts that are invoked within MSPComplete (coming soon). Click [here](https://help.bittitan.com/hc/en-us/articles/115013395648-Writing-and-Using-Automation-Scripts) for more information about writing and using automation scripts. 
* `/Tools (Tools)` - Utility and console tool scripts.

## Initialize-MSPC_Context
Use the Initialize-MSPC_Context cmdlet to create a global mspc context, which contains a number of useful fields:

* `$mspc.Ticket` - Represents an unscoped ticket used for **BT** cmdlets.
* `$mspc.CustomerTicket` - Represents a ticket scoped to a Customer and is used for **BT** cmdlets.
* `$mspc.WorkgroupTicket` - Represents a ticket scoped to a Workgroup and is used for **BT** cmdlets.
* `$mspc.MigrationWizTicket` - Represents a ticket used for **MW** cmdlets.
* `$mspc.Customer` - Represents the Customer retrieved using the id passed to the Init script.
* `$mspc.Workgroup` - Represents the Workgroup retrieved using the id passed to the Init script.
* `$mspc.Data` - Represents hashtable that stores global data.

The $mspc context object created by this cmdlet is identical to the $mspc variable available to every script running on the MSPC automation platform. 
Thus it is recommended to run the Initialize-MSPC_Context cmdlet before testing Runbook scripts locally.
Since the $mspc context created is a global variable, it is useful when debugging a single script or multiple scripts as you do not need to enter credentials and variable inputs each time.

The follow switches can also be used:

```powershell
# Case 1: Initialize an mspc context with a customer ID
# Note: the customer's workgroup is used in creating both the workgroup and workgroup ticket.
Initialize-MSPC_Context -Credentials $credentials -CustomerId 'your customer ID here'

# Case 2: Initialize an mspc context with a workgroup ID
# Note: no customer nor customer ticket are created in this case.
Initialize-MSPC_Context -Credentials $credentials -WorkgroupId 'your workgroup ID here'

# Case 3: Clear the existing global $mspc context
# Note: clears the existing $mspc context before creating a new mspc context.
Initialize-MSPC_Context -Clear
```
More information about [MigrationWiz concepts](https://www.bittitan.com/doc/powershell.html#PagePowerShellmigrationwizmd) and [MSPComplete concepts](https://www.bittitan.com/doc/powershell.html#PagePowerShellmspcmd-overview) exist on the BitTitan website.

## Documentation 
PowerShell documentation can be found on the [BitTitan website](https://www.bittitan.com/doc/powershell.html).