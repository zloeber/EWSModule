function Set-ServerCertificateValidationCallback {
    <#
    .SYNOPSIS
    Sets the current certificate validation callback setting
    .DESCRIPTION
    Sets the current certificate validation callback setting

    .PARAMETER CertCallback
    Defaults to [System.Net.ServicePointManager]::ServerCertificateValidationCallback
     
    .EXAMPLE
    Set-ServerCertificateValidationCallback

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

    param(
        [string]$CertCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
    )

    $script:modCertCallback = $CertCallback
}