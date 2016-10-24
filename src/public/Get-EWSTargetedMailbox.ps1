function Get-EWSTargettedMailbox {
    <#
    .SYNOPSIS
        Return the intended targeted mailbox for ews operations.
    .DESCRIPTION
        Return the intended targeted mailbox for operations. If an email address string is passed we will try to connect to it with non-impersonation rights.
        If the Mailbox parameter is empty or null then we will look at the ews object to see if impersonation is set and return that mailbox if found. Otherwise
        we use the ews object login ID.
    .PARAMETER EWSService
        Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
        Mailbox to target. If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.
    .EXAMPLE
        Get-EWSTargetedMailbox -Mailbox jdoe

        Description
        -----------
        Reterns the email address jdoe from the domain.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox you are targeting.')]
        [string]$Mailbox
    )
    
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    
    if ($EWSService -eq $null) {
        Write-Verbose "$($FunctionName): Using module local ews service object"
        $EWSService = Get-EWSService
        Write-Verbose "$($FunctionName): URL targeted = $($EWSService.URL)"
    }
    
    if ($EWSService -eq $null) {
        throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
    }
    
    if (-not [string]::IsNullOrEmpty($Mailbox)) {
        if (Test-EmailAddressFormat $Mailbox) {
            $email = $Mailbox
        }
        else {
            try {
                $email = Get-EmailAddressFromAD $Mailbox
            }
            catch {
                throw "$($FunctionName): Unable to get a mailbox for this account from AD. Ensure you are running this from a domain joined computer."
            }
        }
    }
    else {
        if ($EWSService.ImpersonatedUserId -ne $null) {
            $impID = $EWSService.ImpersonatedUserId.Id
        }
        else {
            $impID = $EWSService.Credentials.Credentials.UserName
        }
        
        if (-not (Test-EmailAddressFormat $impID)) {
            try {
                $email = ($EWSService.ResolveName("smtp:$($ImpID)@",[ews_resolvenamelocation]::DirectoryOnly, $false)).Mailbox -creplace '(?s)^.*\:', '' -creplace '>',''
            }
            catch {
                throw "$($FunctionName): Unable to find a mailbox with this account."
            }
        }
        else {
            $email = $impID
        }
    }
    
    return $email
}