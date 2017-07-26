<#
.NOTES
	Company:		BitTitan, Inc.
	Title:			GetOutlookVersionReport.PS1
	Author:			SUPPORT@BITTITAN.COM
	Requirements: 
	
	Version:		1.00
	Date:			MARCH 23, 2017

	Exchange Version:	2016
	Windows Version:	WINDOWS 10 ENTERPRISE

	Disclaimer: 	This script is provided ‘AS IS’. No warranty is provided either expresses or implied.

	Copyright: 		Copyright © 2017 BitTitan. All rights reserved.
	
.SYNOPSIS
	Generates a full report of outlook version for a given customer reported by DMA.

.DESCRIPTION 	
	This script retrieves all DMA statistics data for a given customer and generates a report. 

.INPUTS
	-[ManagementProxy.ManagementService.Ticket] Ticket, the ticket for authentication.
    -[guid] CustomerId, the id of the customer to report.
	-[string] Csv, the csv output path.
	-[int] PageSize, the batch size of every retrieve request.
    -[string] Env, the context to work with. Valid options : BT, China.

.EXAMPLE
  	.\GetOutlookVersionReport.ps1 -Ticket $BTTicket -customerId '12345678-0000-0000-0000-000000000000' -csv '.\output.csv' -pageSize 100 -env 'BT' 
	Runs the script and outputs the outlook version statistics for customer with id 12345678-0000-0000-0000-000000000000.
#>
param(
    # Ticket for authentication
    [Parameter(Mandatory=$True)]
    [ManagementProxy.ManagementService.Ticket] $Ticket,
   
	# The id of the customer to work with
	[Parameter(Mandatory=$True)]
    [guid] $CustomerId,

	# The csv output file name
    [Parameter(Mandatory=$False)]
    [string] $Csv = ".\VersionReport-$CustomerId.csv",

	# The batch size of every retrieve request 
    [Parameter(Mandatory=$False)]
    [int] $PageSize = 100,
	
    # The environment to work with
    [Parameter(Mandatory=$False)]
    [ValidateSet("BT", "China")]
    [string] $Env = "BT"
) 

$currentPath = Split-Path -parent $MyInvocation.MyCommand.Definition
Import-Module "$currentPath\BitTitanPowerShell.dll"

################################################################################
# Joins two objects
################################################################################

<#
.Synopsis
   Adds the input properties of the items to output.

.DESCRIPTION
   This is a helper function to the Join-Object which adds the properties of the item to output.

.EXAMPLE
   AddItemProperties $leftItem $leftProperties $output
   AddItemProperties $rightItem $rightProperties $output

.ERRORS
   If the property hash does not have string name or script block expression.
#>
function AddItemProperties($item, $properties, $output)
{
    if($item -ne $null)
    {
        foreach($property in $properties)
        {
            $propertyHash = $property -as [hashtable]
            if($propertyHash -ne $null)
            {
                $hashName=$propertyHash[“name”] -as [string]
                if($hashName -eq $null)
                {
                    throw “there should be a string Name”  
                }
         
                $expression=$propertyHash[“expression”] -as [scriptblock]
                if($expression -eq $null)
                {
                    throw “there should be a ScriptBlock Expression”  
                }
         
                $_=$item
                $expressionValue=& $expression
         
                $output | add-member -MemberType “NoteProperty” -Name $hashName -Value $expressionValue
            }
            else
            {
                # .psobject.Properties allows you to list the properties of any object, also known as “reflection”
                foreach($itemProperty in $item.psobject.Properties)
                {
                    if ($itemProperty.Name -like $property)
                    {
                        $output | add-member -MemberType “NoteProperty” -Name $itemProperty.Name -Value $itemProperty.Value
                    }
                }
            }
        }
    }
}

<#
.Synopsis
   Writes the joined items into output. 

.DESCRIPTION
   This is a helper function to Join-Object which generates output by the given items and properties.

.EXAMPLE
   WriteJoinObjectOutput $leftItem $leftItemMatchInRight $LeftProperties $RightProperties 
#>    
function WriteJoinObjectOutput($leftItem, $rightItem, $leftProperties, $rightProperties)
{
    $output = new-object psobject

    AddItemProperties $leftItem $leftProperties $output
    AddItemProperties $rightItem $rightProperties $output

    $output
}

<#
.Synopsis
   Joins two lists of objects based on the WhereLeftPropertyName equals WhereRightPropertyName. The LeftProperties and RightProperties will also be output.

.DESCRIPTION
   This function will joins two collection of objects based on the where-properties and generate a collection of output with all properties specified.

.EXAMPLE
   Join-Object $a $b "Id" "Id" ("Name","Salary") ("Title","Departement")
#>
function Join-Object
{
    Param
    (
        # List to join with $Right
        [Parameter(Mandatory=$true,
                   Position=0)]
        [object[]]
        $Left,

        # List to join with $Left
        [Parameter(Mandatory=$true,
                   Position=1)]
        [object[]]
        $Right,

        # Condition in which an item in the left matches an item in the right
        [Parameter(Mandatory=$true,
                   Position=2)]
        [string]
        $WhereLeftPropertyName,

        # Condition in which an item in the left matches an item in the right
        [Parameter(Mandatory=$true,
                   Position=3)]
        [string]
        $WhereRightPropertyName,

        # Properties from $Left we want in the output.
        # Each property can:
        # – Be a plain property name like “Name”
        # – Contain wildcards like “*”
        # – Be a hashtable like @{Name=”Product Name”;Expression={$_.Name}}. Name is the output property name
        #   and Expression is the property value. The same syntax is available in select-object and it is 
        #   important for join-object because joined lists could have a property with the same name
        [Parameter(Mandatory=$true,
                   Position=4)]
        [object[]]
        $LeftProperties,

        # Properties from $Right we want in the output.
        # Like LeftProperties, each can be a plain name, wildcard or hashtable. See the LeftProperties comments.
        [Parameter(Mandatory=$true,
                   Position=5)]
        [object[]]
        $RightProperties
    )

    Begin
    {
        # a list of the matches in right for each object in left
        $leftMatchesInRight = new-object System.Collections.ArrayList
    }

    Process
    {
        # build a hash table of the matches
        $hashMap = @{}
        foreach($rightItem in $Right)
        {
           $key = $rightItem.$WhereRightPropertyName
           if (!$hashMap.ContainsKey($key)) 
           {
                $value = New-Object System.Collections.ArrayList
                [void]$value.Add($rightItem)
                [void]$hashMap.Add($key, $value)
           }
           else
           {
                [void]$hashMap[$key].Add($rightItem)
           }
        }
        
        # go over items in $Left and produce the list of matches
        foreach($leftItem in $Left)
        {
            $leftItemMatchesInRight = new-object System.Collections.ArrayList
            $null = $leftMatchesInRight.Add($leftItemMatchesInRight)
            
            $key = $leftItem.$WhereLeftPropertyName
            
            if($hashMap.ContainsKey($key))
            {
                $null = $leftItemMatchesInRight.AddRange($hashMap[$key])               
            }                     
        }

        # go over the list of matches and produce output
        for($i=0; $i -lt $Left.Count;$i++)
        {
            $leftItemMatchesInRight=$leftMatchesInRight[$i]
            $leftItem=$Left[$i]
                               
            if($leftItemMatchesInRight.Count -eq 0)
            {
                continue
            }

            foreach($leftItemMatchInRight in $leftItemMatchesInRight)
            {
                WriteJoinObjectOutput $leftItem $leftItemMatchInRight $LeftProperties $RightProperties 
            }
        }
    }
}

################################################################################
# Generates version report
################################################################################

<#
.Synopsis
   Converts the version into a version report summary item.

.DESCRIPTION
   This function takes in a version number and outputs if the version number is supported. 

.EXAMPLE
   Get-VersionReport 14.0.7173.5000
#>
function Get-VersionReport
{
    [OutputType([string])]

     Param
    (
        # Version number
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()] 
        [string] $Version = $null
    )

    Process
    {
        # Undetected and unrecognized version
        if (!$version) { return "Desktop Deployment requirement is to have Microsoft Outlook 2007, 2010, 2013 or 2016 client installed. Current version could not be detected." } 

        # Convert the version into a System.Version
        $systemVersion = New-Object System.Version($version)

        # Output string
        $output = New-Object System.Text.StringBuilder

        # Switch by the Major Version
        switch ($systemVersion.Major)
        {
            11 
            {
                # Not supported check Outlook 2003
                [void]$output.Append("Current version detected: Outlook 2003. The current version of Outlook is not supported. Desktop Deployment requirement is to have Microsoft Outlook 2007, 2010, 2013 or 2016 client installed.")
            }
            12
            {
                # Current version 2007
                [void]$output.Append("Current version detected: Outlook 2007. ")

                # Not supported check Outlook 2007
                if($systemVersion.Build -lt 6665) { [void]$output.Append("The current version of Outlook is not supported. ") }

                # MS Office 2007 SP3
                if ($systemVersion.Build -lt 6607) { [void]$output.Append("MS Office 2007 Service Pack 3 (SP3) installation is required http://www.microsoft.com/en-us/download/details.aspx?id=27838.") }

                # KB2687404
                if ($systemVersion.Build -lt 6665) { [void]$output.Append("KB2687404 update installation is required http://www.microsoft.com/en-us/download/details.aspx?id=35718.") }
                else { [void]$output.Append("The current version of Outlook is supported.") }
            }
            14
            {
                # Current version 2010
                [void]$output.Append("Current version detected: Outlook 2010. ")

                # Not supported check Outlook 2010
                if ($systemVersion.Build -lt 6126) { [void]$output.Append("The current version of Outlook is not supported. ") }

                # Office 2010 SP1 x86 and x64
                if ($systemVersion.Build -lt 6025) { [void]$output.Append("MS Office 2010 Service Pack 1 (SP1) installation is required http://www.microsoft.com/en-us/download/details.aspx?id=26622 for 32-bit, http://www.microsoft.com/en-us/download/details.aspx?id=26617 for 64-bit.") }

                # KB2687623 for x86 and x64
                if ($systemVersion.Build -lt 6126) { [void]$output.Append("KB2687623 update installation is required http://www.microsoft.com/en-us/download/details.aspx?id=35702 for 32-bit, http://www.microsoft.com/en-us/download/details.aspx?id=35714 for 64-bit.") }
                else { [void]$output.Append("The current version of Outlook is supported.") }
            }
            15
            {
                # Current version 2013
                [void]$output.Append("Current version detected: Outlook 2013. ")

                # Check the version of  Outlook 2013 that is supported
                if ($systemVersion.Build -lt 4420) { [void]$output.Append("The current version of Outlook is not supported. Please install the latest Outlook 2013 updates.") }
                else { [void]$output.Append("The current version of Outlook is supported.") }
            }
            16
            {
                # Current version 2016 supported
                [void]$output.Append("Current version detected: Outlook 2016. The current version of Outlook is supported.")
            }
            default
            {
                [void]$output.Append("Desktop Deployment requirement is to have Microsoft Outlook 2007, 2010, 2013 or 2016 client installed. Current version could not be detected.")
            }
        }

        return $output.ToString()
    }
}

################################################################################
# Script to fetch and generate version report
################################################################################

# Retrieve the customer
$customer = Get-BT_Customer -Ticket $Ticket -Environment $Env -FilterBy_Guid_Id $CustomerId
$organizationId = $customer.OrganizationId
if (-not $organizationId) {
    Write-Error "Customer does not exist for $CustomerId, aborted."
    return
}

# Bind the organization Id to the ticket
if (-not $Ticket.IsPrivileged) 
{
    $Ticket = Get-BT_Ticket -Ticket $Ticket -OrganizationId $organizationId -Environment $Env
    Write-Verbose "Currently non-privilege is used, the page size is limited to 100."
    $PageSize = 100
}

# Log Checkpoint
Write-Verbose "Version report will be generated to $Csv for customer $($customer.CompanyName)."

# Log Checkpoint
Write-Verbose "Retrieving all entities..."

# Retrieve the customer device users with pagination
Write-Verbose "Retrieving all customerDeviceUsers...This may take a while"
$count = 0
$customerDeviceUsers = New-Object System.Collections.ArrayList
While($true)
{    
    [array]$temp = Get-BT_CustomerDeviceUser -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Boolean_IsDeleted $False -PageOffset $($count*$PageSize) -PageSize $PageSize
    $customerDeviceUsers.AddRange($temp)
    if ($customerDeviceUsers.Count % 1000 -eq 0) { Write-Verbose "$($customerDeviceUsers.Count) entities retrieved." }
    if ($temp.count -lt $PageSize) { break } 
    $count++
}
Write-Verbose "Totally $($customerDeviceUsers.Count) customerDeviceUsers retrieved."

# Retrieve the customer end users with pagination
Write-Verbose "Retrieving all customerEndUsers...This may take a while"
$count = 0
$customerEndUsers = New-Object System.Collections.ArrayList
While($true)
{    
    [array]$temp = Get-BT_CustomerEndUser -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Boolean_IsDeleted $False -PageOffset $($count*$PageSize) -PageSize $PageSize
    $customerEndUsers.AddRange($temp)
    if ($customerEndUsers.Count % 1000 -eq 0) { Write-Verbose "$($customerEndUsers.Count) entities retrieved." }
    if ($temp.count -lt $PageSize) { break } 
    $count++
}
Write-Verbose "Totally $($customerEndUsers.Count) customerEndUsers retrieved."

# Retrieve the customer devices with pagination
Write-Verbose "Retrieving all customerDevices...This may take a while"
$count = 0
$customerDevices = New-Object System.Collections.ArrayList
While($true)
{   
    [array]$temp = Get-BT_CustomerDevice -Ticket $Ticket -Environment $Env -FilterBy_Guid_OrganizationId $organizationId -FilterBy_Boolean_IsDeleted $False -PageOffset $($count*$PageSize) -PageSize $PageSize
    $customerDevices.AddRange($temp)
    if ($customerDevices.Count % 1000 -eq 0) { Write-Verbose "$($customerDevices.Count) entities retrieved." }
    if ($temp.count -lt $PageSize) { break } 
    $count++   
}
Write-Verbose "Totally $($customerDevices.Count) customerDevices retrieved."

# Log Checkpoint
Write-Verbose "Joining CustomerDeviceUsers and customerEndUsers..."

# First Join Customer Device Users with Customer End Users
$tempObjects = Join-Object -Left $customerDeviceUsers -Right $customerEndUsers -LeftProperties EndUserId,DeviceId,OfficeVersions -RightProperties PrimaryIdentity -WhereLeftPropertyName EndUserId -WhereRightPropertyName Id

# Log Checkpoint
Write-Verbose "Joining customerDevices..."

# Second Join with customer devices
$outputObjects = Join-Object -Left $tempObjects -Right $customerDevices -LeftProperties EndUserId,DeviceId,PrimaryIdentity,OfficeVersions -RightProperties DeviceName -WhereLeftPropertyName DeviceId -WhereRightPropertyName Id

# Log Checkpoint
Write-Verbose "Getting the version infos..."

# Check the customer device user software versions
foreach ($object in $outputObjects)
{           
    # Try parse the version Numbers    
    if ($object.OfficeVersions) 
    {                 
        $versionNumbers = $object.officeVersions.Split(",") 
        
        # If the object is with specified office versions
        if ($versionNumbers[0].Split(".").Count -eq 4) 
        { 
            $versions = $versionNumbers 
        }
    }     

    # If nothing found then clears the officeVersion field
    if (!$versions)
    {
        $object.OfficeVersions = $null                  
    }
    # Get the highest version number for the object  
    else
    {        
        # Convert the version into System.Version to compare
        $latestVersion = New-Object System.Version($versions[0])
        for ($i=1; $i -lt $versions.Count; $i++)
        {
            $tempVersion = New-Object System.Version($versions[$i])
            if ($latestVersion.CompareTo($tempVersion) -eq -1) { $latestVersion = $tempVersion }            
        }

        # Copy the latest version to current object
        $object.OfficeVersions = $latestVersion.ToString()
    }       
}

# Log Checkpoint
Write-Verbose "Outputing csv..."

# Iterate through the $outputObjects and write report
$outputObjects | Select DeviceName,PrimaryIdentity,OfficeVersions, @{ Name = 'Report'; Expression = { Get-VersionReport -Version $_.OfficeVersions } } | Export-Csv $Csv