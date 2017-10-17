# BitTitan PowerShell
This folder contains a set of tools, building blocks and full scenario scripts tailored to work with the BitTitan PowerShell SDK. 

## Quick Startup
1. Install the [BitTitan PowerShell SDK](https://www.bittitan.com/downloads/bittitanpowershellsetup.msi).
2. Clone the [bittitan-sdk repository](https://github.com/BitTitan/bittitan-sdk) or download the desired scripts.
    ```shell
    git clone https://github.com/BitTitan/bittitan-sdk.git
    ```
3. Move the `Init.ps1` (bittitan-sdk/PowerShell/Tools/Init.ps1) script to the working directory. If the repository was not cloned, download the [Init script](https://github.com/BitTitan/bittitan-sdk/blob/master/PowerShell/Tools/Init.ps1) and move it to the working directory.
4. Execute desired scripts from the working directory.

## Contents
* `/BuildingBlocks (Building Blocks)` - Samples to help developers build scripts that automate various operations.
* `/Support (Support Scripts)` - Scripts that enable developers address numerous scenarios by utilizing them.
* `/TaskLibrary (Task Library)` - Scripts that are invoked within MSPComplete.
* `/Tools (Tools)` - Utility and console tool scripts.

## Init Script
All scripts within the BuildingBlocks folder utilize the Init.ps1 script. The script initializes important variables used when invoking BitTitan cmdlets.
Invoking the script is as simple as follows:

```powershell
.\Init.ps1
```

After invoking the script, it will ask for the following:
* The credentials of the BitTitan account.
* The id of the Workgroup being associated to the ticket. For more information about Workgroups, please see [MSPComplete concepts](https://www.bittitan.com/doc/powershell.html#PagePowerShellmspcmd-overview).

The follow switches can also be used:

```powershell
# Clears the global variables $mspc and $creds
.\Init.ps1 -Clear

# Prompts for the Customer id and initializes customer related variables
.\Init.ps1 -Customer
```

Initalized Variables
* `$mspc.Ticket` - Represents an unscoped ticket used for **BT** cmdlets.
* `$mspc.CustomerTicket` - Represents a ticket scoped to a Customer and is used for **BT** cmdlets.
* `$mspc.WorkgroupTicket` - Represents a ticket scoped to a Workgroup and is used for **BT** cmdlets.
* `$mspc.MigrationWizTicket` - Represents a ticket used for **MW** cmdlets.
* `$mspc.Customer` - Represents the Customer retrieved using the id passed to the Init script.
* `$mspc.Workgroup` - Represents the Workgroup retrieved using the id passed to the Init script.

More information about [MigrationWiz concepts](https://www.bittitan.com/doc/powershell.html#PagePowerShellmigrationwizmd) and [MSPComplete concepts](https://www.bittitan.com/doc/powershell.html#PagePowerShellmspcmd-overview) exist on the BitTitan website.

## Documentation 
PowerShell documentation can be found on the [BitTitan website](https://www.bittitan.com/doc/powershell.html).