function Set-EWSMailboxImpersonation {
    <# 
    .SYNOPSIS 
    Set the impersonation for a mailbox.
    .DESCRIPTION 
    Set the impersonation for a mailbox.
    .PARAMETER EWSService
    Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
    Mailbox to target.
    .PARAMETER SkipValidation
    Do not validate if you have impersonation rights for the mailbox (can speed things up quite a bit)

    .EXAMPLE
    Set-EWSMailboxImpersonation -Mailbox mailbox@domain.com

    Set impersonation mode for the current connected EWS user for mailbox@domain.com

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
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(Position=1, Mandatory=$True, HelpMessage='Mailbox to impersonate.')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Do not attempt to validate rights against this mailbox (can speed up operations)')]
        [switch]$SkipValidation
    )
  
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand

    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    
    if (($EWSService -eq $null) -and ((Get-EWSService) -ne $null)) {
        Write-Verbose "$($FunctionName): Using module local ews service object"
        $EWSService = Get-EWSService
        Write-Verbose "$($FunctionName): URL targeted = $($EWSService.URL)"
    }
    else {
        throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
    }

    if (Test-EmailAddressFormat $Mailbox) {
        $enumType = [ews_connidtype]::SmtpAddress
    }
    else {
        $enumType = [ews_connidtype]::PrincipalName
    }
    try {
        $EWSService.ImpersonatedUserId = New-Object ews_impersonateuserid($enumType,$Mailbox)
        if (-not $SkipValidation) {
            $InboxFolder= new-object ews_folderid([ews_wellknownfolder]::Inbox,$Mailbox)
            $Inbox = [ews_folder]::Bind($EWSService,$InboxFolder)
        }
    }
    catch {
        Write-Error ('Set-EWSMailboxImpersonation: Unable to impersonate {0}, check to see that you have adequately assigned permissions to impersonate this account.' -f $Mailbox)
        throw $_.Exception.Message  
    }
}