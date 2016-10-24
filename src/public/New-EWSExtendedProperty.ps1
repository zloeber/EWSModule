function New-EWSExtendedProperty {
    <# 
    .SYNOPSIS 
        Creates a new extended property which can be assigned to items in outlook.
    .DESCRIPTION 
        Creates a new extended property which can be assigned to items in outlook. These are generally hidden to end 
        users but can be invaluable in creating items that you can then later locate again and know they were created by
        your processes.
    .PARAMETER PropertyType
        Type of extended property to create.
    .PARAMETER PropertyName
        Name of extended property
    .EXAMPLE
        $CalEntry = New-EWSCalendarEntry -IsAllDayEvent $true -Subject 'My Event' -Location 'Elsewhere'

        Description
        --------------
        Creates a new calendar entry as an all day event called 'My Event' and stores it in $CalEntry

    .NOTES
        Author: Zachary Loeber
        Site: http://www.the-little-things.net/
        Requires: Powershell 3.0
        Version History
        1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Type of extended property to create.')]
        [ValidateNotNullOrEmpty()]
        [ews_mapiproptype[]]$PropertyType = [System.Enum]::GetValues([ews_mapiproptype]),
        [parameter(Position=1, Mandatory=$True, HelpMessage='Name of extended property')]
        [ValidateNotNullOrEmpty()]
        [string]$PropertyName
    )
    
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand

    if (-not (Get-EWSModuleInitializationState)) {
        throw "$(FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    if ($PropertyType.Count -gt 1) {
        $PropertyType = [ews_mapiproptype]::String
    }
    Write-Verbose "$(FunctionName): Attempting to create an extended property"
    return New-Object -TypeName ews_extendedpropdef([ews_extendedpropset]::PublicStrings, $PropertyName, $PropertyType)
}