function Set-EWSOofSettings {
     <# 
    .SYNOPSIS 
        Set the out of office settings for a mailbox.
    .DESCRIPTION 
        Set the out of office settings for a mailbox.
    .PARAMETER EWSService
        Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
        Mailbox to target.
    .PARAMETER State
        State of OOF for the mailbox. Can be Enabled, Disabled, or Scheduled. Defaults to Disabled.
    .PARAMETER ExternalAudience
        Whom will get OOF externally. Can be All, Known, or None. Defaults to All.
    .PARAMETER StartTime
        Start time that OOF replies will be scheduled.
    .PARAMETER EndTime
        End time that OOF replies will be enabled or scheduled.
    .PARAMETER InternalReply
        Internal OOF message.
    .PARAMETER ExternalReply
        External OOF message.

    .EXAMPLE
        Set-EWSOOFSettings -Mailbox mailbox@domain.com

        Description
        --------------
        Disables the OOF settings for mailbox@domain.com

    .NOTES
        Author: Zachary Loeber
        Site: http://www.the-little-things.net/
        Requires: Powershell 3.0
        Version History
        1.0.0 - Initial release
    #>
    [CmdletBinding(DefaultParameterSetName='Default')]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.', ParameterSetName = 'Default')]
        [parameter(Position=0, HelpMessage='Connected EWS object.', ParameterSetName = 'Enabled')]
        [parameter(Position=0, HelpMessage='Connected EWS object.', ParameterSetName = 'Disabled')]
        [parameter(Position=0, HelpMessage='Connected EWS object.', ParameterSetName = 'Scheduled')]
        [ews_service]$EWSService,
        [parameter(Position=1, Mandatory=$True, ParameterSetName = 'Default', ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [parameter(Position=1, Mandatory=$True, ParameterSetName = 'Enabled', ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [parameter(Position=1, Mandatory=$True, ParameterSetName = 'Disabled', ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [parameter(Position=1, Mandatory=$True, ParameterSetName = 'Scheduled', ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [String]$Mailbox,
        [Parameter(Position=2, ParameterSetName = 'Default')]
        [Parameter(Position=2, ParameterSetName = 'Enabled')]
        [Parameter(Position=2, ParameterSetName = 'Disabled')]
        [Parameter(Position=2, ParameterSetName = 'Scheduled')]
        [ValidateSet("Enabled","Disabled","Scheduled")]
        [String]$State = 'Disabled',
        [Parameter(Position=3, ParameterSetName = 'Enabled')]
        [Parameter(Position=3, ParameterSetName = 'Scheduled')]
        [ValidateSet("All","External","None")]
        [String]$ExternalAudience = 'All',
        [Parameter(Position=4, ParameterSetName = 'Enabled')]
        [Parameter(Position=4, ParameterSetName = 'Scheduled')]
        [DateTime]$StartTime,
        [Parameter(Position=5, ParameterSetName = 'Enabled')]
        [Parameter(Position=5, ParameterSetName = 'Scheduled')]
        [DateTime]$EndTime,        
        [Parameter(Position=6, ParameterSetName = 'Enabled')]
        [Parameter(Position=6, ParameterSetName = 'Scheduled')]
        [String]$InternalReply,
        [Parameter(Position=7, ParameterSetName = 'Enabled')]
        [Parameter(Position=7, ParameterSetName = 'Scheduled')]
        [String]$ExternalReply
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
    }
    catch {
        throw "$($FunctionName): Unable to get oof settings from $TargetMailbox"
    }

    if($StartTime -and $EndTime) {
        $Duration = New-Object Microsoft.Exchange.WebServices.Data.TimeWindow -arg $StartTime,$EndTime
        $PSBoundParameters.Duration = $Duration
        $PSBoundParameters.State = "Scheduled"
        [Void]$PSBoundParameters.remove("StartTime")
        [Void]$PSBoundParameters.remove("EndTime")
    }
    
    foreach($p in $PSBoundParameters.GetEnumerator()) {
        if (($p.key -ne "Mailbox") -and ($p.key -ne "EWSService"))  {
            $oof."$($p.key)" = $p.value
        }
    }

    $oof.State = [Microsoft.Exchange.WebServices.Data.OofState]::$State
    $EWSService.SetUserOofSettings($Mailbox,$oof)
}