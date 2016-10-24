function Get-EWSOOFSettings {
    <# 
    .SYNOPSIS 
        Get the out of office settings for a mailbox.
    .DESCRIPTION 
        Get the out of office settings for a mailbox.
    .PARAMETER EWSService
        Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
        Mailbox to target.

    .EXAMPLE
        Get-EWSOOFSettings -Mailbox mailbox@domain.com

        Description
        --------------
        Get the out of office settings for mailbox@domain.com

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
        [parameter(Position=1, Mandatory=$True, ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [String]$Mailbox
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
    try {
        $TargetedMailbox = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
    }
    catch {
        throw "$($FunctionName): Unable to target $Mailbox"
    }
    try {
        $oof = $EWSService.GetUserOofSettings($TargetedMailbox)
        New-Object PSObject -Property @{
            State = $oof.State
            ExternalAudience = $oof.ExternalAudience
            StartTime = $oof.Duration.StartTime
            EndTime = $oof.Duration.EndTime
            InternalReply = $oof.InternalReply
            ExternalReply = $oof.ExternalReply
            AllowExternalOof = $oof.AllowExternalOof
            Mailbox = $TargetedMailbox
        }
    }
    catch {
        throw "$($FunctionName): Unable to get out of office info for $TargetedMailbox"
    }
}