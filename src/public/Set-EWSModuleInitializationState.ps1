function Set-EWSModuleInitializationState {
     <#
    .SYNOPSIS
    Sets the module initialization state
    .DESCRIPTION
    Sets the module initialization state
    .PARAMETER State
    State of the module initialization, either true or fals
    .EXAMPLE
    Set-EWSModuleInitializationState $true          

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    param(
        [bool]$State = $false
    )
    $script:EWSModuleInitialized = $State
}