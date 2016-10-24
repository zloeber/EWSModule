function Get-ServerCertificateValidationCallback {
    <#
    .SYNOPSIS
    Returns the current certificate validation callback setting
    .DESCRIPTION
    Returns the current certificate validation callback setting
     
    .EXAMPLE
    Get-ServerCertificateValidationCallback         

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    return $script:modCertCallback
}