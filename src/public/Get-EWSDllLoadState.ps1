function Get-EWSDllLoadState {
    if (-not (get-module Microsoft.Exchange.WebServices)) {
        return $false
    }
    else {
        return $true
    }
}