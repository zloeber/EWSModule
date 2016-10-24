function Import-EWSDll {
    <#
    .SYNOPSIS
    Load EWS dlls.

    .DESCRIPTION
    Load EWS dlls.
     
    .PARAMETER EWSManagedApiPath
    Full path to Microsoft.Exchange.WebServices.dll. If not provided we will try to load it from several best guess locations.
    
    .EXAMPLE
    Import-EWSDll
         
    .NOTES
    This function requires Exchange Web Services Managed API. From what I can tell you don't even need to install the msi. AS long
    as the Microsoft.Exchange.WebServices.dll file is extracted and available that should work.
    
    The EWS Managed API can be obtained from: http://www.microsoft.com/en-us/download/details.aspx?id=28952    

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=0)]
        [string]$EWSManagedApiPath
    )
    
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand
    
    $ewspaths = @()
    if (-not (Get-EWSDllLoadState)) {
        if (-not [string]::IsNullOrEmpty($EWSManagedApiPath)) {
            $ewspaths += @($EWSManagedApiPath)
        }
        $ewspaths += $script:ewsdllpaths

        $EWSLoaded = $false
        foreach ($ewspath in $ewspaths) {
            try {
                if (-not $EWSLoaded) {
                    if (Test-Path $ewspath) {
                        Write-Verbose "$($FunctionName): Attempting to load $ewspath"
                        Import-Module -Name $ewspath -ErrorAction:Stop -Global
                        $EWSLoaded = $true
                    }
                }
            }
            catch {}
        }
    }
    else {
        Write-Verbose ("$($FunctionName): EWS dll already Loaded!")
    }
}