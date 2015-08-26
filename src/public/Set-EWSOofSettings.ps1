function Set-EWSOofSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [String]$Identity,
        [Parameter(Position=1)]
        [String]$State,
        [Parameter(Position=2)]
        [String]$ExternalAudience,
        [Parameter(Position=3)]
        [DateTime]$StartTime,
        [Parameter(Position=4)]
        [DateTime]$EndTime,        
        [Parameter(Position=5)]
        [String]$InternalReply,
        [Parameter(Position=6)]
        [String]$ExternalReply
    )
    
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
        if($p.key -ne "Identity") {
            $oof."$($p.key)" = $p.value                
        }
    }
    $EWSService.SetUserOofSettings($Identity,$oof)
}