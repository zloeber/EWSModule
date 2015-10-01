function Get-EWSUserDN {
    [CmdletBinding()] 
    param(
        [parameter(HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [Parameter(Position=1, Mandatory=$true)]
        [string]$EmailAddress
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
    $ExchangeVersion = [ews_exchver]::$($EWSService.RequestedServerVersion)
    $adService = New-Object ews_autod($ExchangeVersion)
    $adService.Credentials = $EWSService.Credentials
    $adService.EnableScpLookup = $false;
    $adService.RedirectionUrlValidationCallback = {$true}
    $UserSettings = new-object ews_usersettingname[] 1
    $UserSettings[0] = [ews_usersettingname]::UserDN
    $adResponse = $adService.GetUserSettings($EmailAddress, $UserSettings);
    return $adResponse.Settings[[ews_usersettingname]::UserDN]
}