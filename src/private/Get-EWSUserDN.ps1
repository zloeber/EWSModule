function Get-EWSUserDN {
    param (
            [Parameter(Position=0, Mandatory=$true)] [string]$EmailAddress,
            [Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
          )
    process{
        $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
        $adService = New-Object Microsoft.Exchange.WebServices.AutoDiscover.AutodiscoverService($ExchangeVersion);
        $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
        $adService.Credentials = $creds
        $adService.EnableScpLookup = $false;
        $adService.RedirectionUrlValidationCallback = {$true}
        $UserSettings = new-object Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[] 1
        $UserSettings[0] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN
        $adResponse = $adService.GetUserSettings($EmailAddress , $UserSettings);
        return $adResponse.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN]
    }
}