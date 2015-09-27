function Get-EWSAutoDiscoverPhotoURL{
    param (
        $EmailAddress="$( throw 'Email is a mandatory Parameter' )",
        $Credentials="$( throw 'Credentials is a mandatory Parameter' )"
    )
    $version= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
    $adService= New-Object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($version);
    $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
    $adService.Credentials = $creds
    $adService.EnableScpLookup=$false;
    $adService.RedirectionUrlValidationCallback= {$true}
    $adService.PreAuthenticate=$true;
    $UserSettings= new-object Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[] 1
    $UserSettings[0] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::ExternalPhotosUrl
    $adResponse=$adService.GetUserSettings($EmailAddress, $UserSettings)
    $PhotoURI= $adResponse.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::ExternalPhotosUrl]
    return $PhotoURI.ToString()
}