# This PowerShell Script includes three functions to writes a log File
# Functions:
# Log-Write       Appends a New Line to the end of the log file; if the log does not exist, it will be genereated
# Log-Error       Writes the passed error to a new line at the end of the specified log file
# Log-Finish      Writes finishing logging data to specified log and then exits the calling script
#
# to inliude this logging features in other PowerShell Scripts, copy thhs script into the same folder to the Calling Script
# and include the following lines into the calling script
#-> including a separate Function-File for logging
#-> $directorypath = Split-Path $MyInvocation.MyCommand.Path 
#-> $incFunctions = $directorypath + "\Logging_Functions.ps1"  
#-> $logfile = $directorypath + "\Migration.log"                            #The name of the log file, may be changed
#-> . $incFunctions                                                         # use the . syntax to include the functions file 
#


Function Log-Write{
  <#
  .SYNOPSIS
    Writes to a log file

  .DESCRIPTION
    Appends a new line to the end of the specified log file
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Migration.log
  
  .PARAMETER LineValue
    Mandatory. The string that you want to write to the log
      
  .INPUTS
    Parameters above

  .OUTPUTS
    None

  .NOTES
 
    Version:        1.0
    Author:         Hans Brender
    Creation Date:  12/25/15
    Purpose/Change: Initial function development

  .EXAMPLE
    Log-Write -LogPath "C:\Windows\Temp\Migration.log" -LineValue "This is a new line which I am appending to the end of the log file."
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$LogPath, [Parameter(Mandatory=$true)][string]$LineValue)
  
  Process{
        $out = "$([DateTime]::Now) :" + $LineValue
        Add-Content -Path $LogPath -Value $Out
  
    #Write to screen for debug mode
    Write-Debug $LineValue
  }
}

Function Log-Error{
  <#
  .SYNOPSIS
    Writes an error to a log file

  .DESCRIPTION
    Writes the passed error to a new line at the end of the specified log file
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Migration.log
  
  .PARAMETER ErrorDesc
    Mandatory. The description of the error you want to pass (use $_.Exception)
  
  .PARAMETER ExitGracefully
    Mandatory. Boolean. If set to True, runs Log-Finish and then exits script

  .INPUTS
    Parameters above

  .OUTPUTS
    None

  .NOTES
    Version:        1.0
    Author:         Hans Brender
    Creation Date:  12/25/12
    Purpose/Change: Initial function development
    
  .EXAMPLE
    Log-Error -LogPath "C:\Windows\Temp\Migration.log" -ErrorDesc $_.Exception -ExitGracefully $True
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$LogPath, [Parameter(Mandatory=$true)][string]$ErrorDesc, [Parameter(Mandatory=$true)][boolean]$ExitGracefully)
  
  Process{
        Add-Content -Path $LogPath -Value "$([DateTime]::Now) :Error [$ErrorDesc]."
  
    #Write to screen for debug mode
    Write-Debug "Error  [$ErrorDesc]."
    
    #If $ExitGracefully = True then run Log-Finish and exit script
    If ($ExitGracefully -eq $True){
      Log-Finish -LogPath $LogPath
      Break
    }
  }
}

Function Log-Finish{
  <#
  .SYNOPSIS
    Write closing logging data & exit

  .DESCRIPTION
    Writes finishing logging data to specified log and then exits the calling script
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write finishing data to. Example: C:\Windows\Temp\Migration.log

  .PARAMETER NoExit
    Optional. If this is set to True, then the function will not exit the calling script, so that further execution can occur
  
  .INPUTS
    Parameters above

  .OUTPUTS
    None

  .NOTES
    Version:        1.0
    Author:         Hans Brender
    Creation Date:  12/25/15
    Purpose/Change: Initial function development
    

  .EXAMPLE
    Log-Finish -LogPath "C:\Windows\Temp\Migration.log"

  .EXAMPLE
    Log-Finish -LogPath "C:\Windows\Temp\Migration.log" -NoExit $True
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$LogPath, [Parameter(Mandatory=$false)][string]$NoExit)
  
  Process{
    Add-Content -Path $LogPath -Value "$([DateTime]::Now) :Finished processing."
  
    #Write to screen for debug mode
    Write-Debug ""
    Write-Debug "Finished processing at [$([DateTime]::Now)]."

  
    #Exit calling script if NoExit has not been specified or is set to False
    If(!($NoExit) -or ($NoExit -eq $False)){
      Exit
    }    
  }
}