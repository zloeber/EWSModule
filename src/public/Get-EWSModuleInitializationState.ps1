function Get-EWSModuleInitializationState {
    <#
    .SYNOPSIS
        Returns the initialization state of the module.
    .DESCRIPTION
        Returns the initialization state of the module.

    .EXAMPLE
        Get-EWSService

        Description
        --------------
        Returns the initialization state of the module.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    return $script:EWSModuleInitialized
}