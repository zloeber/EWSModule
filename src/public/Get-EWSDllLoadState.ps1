function Get-EWSDllLoadState {
    <#
    .SYNOPSIS
    Determine if the EWS dll is loaded or not.
    .DESCRIPTION
    Determine if the EWS dll is loaded or not.
    .EXAMPLE
    Get-EWSDllLoadState

    Determine if the EWS dll is loaded or not.
    .LINK
    http://www.the-little-things.net/
    .LINK
    https://www.github.com/zloeber/EWSModule
    .NOTES
    Author: Zachary Loeber
    Requires: Powershell 3.0
    Version History
    1.0.0 - Initial release
    #>
    if (get-module Microsoft.Exchange.WebServices) {
        return $true
    }
    else {
        return $false
    }
}