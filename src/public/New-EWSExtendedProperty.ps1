function New-EWSExtendedProperty {
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Type of extended property to create.')]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Exchange.WebServices.Data.MapiPropertyType[]]$PropertyType = [System.Enum]::GetValues([Microsoft.Exchange.WebServices.Data.MapiPropertyType]),
        [parameter(Position=1, Mandatory=$True, HelpMessage='Name of extended property')]
        [ValidateNotNullOrEmpty()]
        [string]$PropertyName
    )
    if (-not (Get-EWSModuleInitializationState)) {
        throw 'EWS Module has not been initialized. Try running Initialize-EWS to rectify.'
    }
    if ($PropertyType.Count -gt 1) {
        $PropertyType = [ews_mapiproptype]::String
    }
    Write-Verbose "New-EWSExtendedProperty: Attempting to create an extended property"
    return New-Object -TypeName ews_extendedpropdef([ews_extendedpropset]::PublicStrings, $PropertyName, $PropertyType)
}