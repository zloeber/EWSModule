function Get-EWSService {
     <#
    .SYNOPSIS
        Returns the current EWSService module variable object
    .DESCRIPTION
        Returns the current EWSService module variable object

    .EXAMPLE
        Get-EWSService

        Description
        --------------
        Returns the current EWSService module variable object

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    return $script:modEWSService
}