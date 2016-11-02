function Get-EWSTemplateFunction {
    <#
    .SYNOPSIS
    .DESCRIPTION
    .PARAMETER EWSService
    Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
    Mailbox to target. If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.

    .EXAMPLE
    PS > Get-EWSTemplateFunction -Mailbox jdoe@contoso.com

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
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox of folder.')]
        [string]$Mailbox
    )
    # Pull in all the caller verbose,debug,info,warn and other preferences
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    
    if ($EWSService -eq $null) {
        Write-Verbose "$($FunctionName): Using module local ews service object"
        $EWSService = Get-EWSService
    }
    
    if ($EWSService -eq $null) {
        throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
    }

    try {
        $TargetedMailbox = Get-EWSTargetedMailbox -EWSService $EWSService -Mailbox $Mailbox
        Write-Verbose "$($FunctionName): Targeted Mailbox = $($TargetedMailbox) "
        $ConnectedMB = New-Object ews_mailbox($TargetedMailbox)
    }
    catch {
        throw "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
    }

    # Start function here
}