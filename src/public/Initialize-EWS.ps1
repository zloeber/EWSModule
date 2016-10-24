function Initialize-EWS {
    <#
    .SYNOPSIS
    Load EWS dlls and create type accelerators for other functions.

    .DESCRIPTION
    Load EWS dlls and create type accelerators for other functions.
     
    .PARAMETER EWSManagedApiPath
    Full path to Microsoft.Exchange.WebServices.dll. If not provided we will try to load it from several best guess locations.
    
    .PARAMETER Uninitialize
    Remove previously added type-accelerators.
     
    .EXAMPLE
    Initialize-EWS
         
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
        [string]$EWSManagedApiPath,
        [parameter(Position=1)]
        [switch]$Uninitialize
    )

    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not $Uninitialize) {
        Import-EWSDll -EWSManagedApiPath $EWSManagedApiPath
        if (Get-EWSDllLoadState) {
            if (-not (Get-EWSModuleInitializationState)) {
                # Setup a bunch of type accelerators to make this mess easier to understand (slightly)
                $accelerators = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')
                
                Add-Type -AssemblyName Microsoft.Exchange.WebServices
                Write-Verbose ("$($FunctionName): Attempting to create type accelerators.")
                foreach ($Key in ($script:EWSAccels).Keys) {
                    Write-Verbose "$($FunctionName): Adding type accelerator - $Key for the type $($Script:EWSAccels[$Key])"
                    $accelerators::Add($Key,$script:EWSAccels[$Key])
                }
                
                # Powershell 5.0 needs this or nothing will work (dammit!)
                if ($PSVersionTable.PSVersion.Major -eq 5) {
                    $builtinfield = $accelerators.GetField('builtinTypeAccelerators',[System.Reflection.BindingFlags]'Static,NonPublic')
                    $builtinfield.SetValue($builtinfield,$accelerators::Get)
                }

                Set-EWSModuleInitializationState $true
                return $true
            }
            else {
                return $true
            }
        }
        else {
            throw "$($FunctionName): Cant load EWS module. Please verify it is installed or manually provide the path to Microsoft.Exchange.WebServices.dll"
        }
    }
    else {
        # Uninitialize EWS
        if (Get-EWSModuleInitializationState) {
            $accelerators = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')::get
            $accelkeyscopy = @{}
            $accelerators.Keys | Where {$_ -like "ews_*"} | Foreach { $accelkeyscopy.$_ = $EWSAccels[$_] }
            foreach ( $key in $accelkeyscopy.Keys ) {
                Write-Verbose "UnInitialize-EWS: Removing type accelerator - $($key)"
                $accelerators.Remove($key) | Out-Null
            }
            Write-Verbose ("$($FunctionName): Custom type accelerators removed!")
            Set-EWSModuleInitializationState $false
        }
        if (Get-EWSDllLoadState) {
            Remove-Module Microsoft.Exchange.WebServices
            Write-Verbose ("$($FunctionName): EWS dll Unloaded!")
        }

        return $true
    }
}