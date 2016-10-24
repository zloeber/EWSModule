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

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>

    param(
        [string]$CertCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
    )

    $script:modCertCallback = $CertCallback
}