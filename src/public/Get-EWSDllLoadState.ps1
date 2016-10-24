function Get-EWSDllLoadState {
   <#
    .SYNOPSIS
        Determine if the EWS dll is loaded or not.
    .DESCRIPTION
        Determine if the EWS dll is loaded or not.

    .EXAMPLE
        Get-EWSDllLoadState

        Description
        --------------
        Determine if the EWS dll is loaded or not.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    if (-not (get-module Microsoft.Exchange.WebServices)) {
        return $false
    }
    else {
        return $true
    }
}