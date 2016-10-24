function Set-EWSService {
    <#
    .SYNOPSIS
    Sets the module connected service.
    .DESCRIPTION
    Sets the module connected service.
    .PARAMETER ConnectedService
    Connected service to set for module. Defaults to null
     
    .EXAMPLE
    Set-EWSService        

    .NOTES
    Author: Zachary Loeber
    Site: http://www.the-little-things.net/
    Requires: Powershell 3.0
    Version History
    1.0.0 - Initial release
    #>
    param(
        [ews_service]$ConnectedService = $null
    )

    $script:modEWSService = $ConnectedService
}