function New-EWSCalendarEntry {
    # Returns an appointment to be manipulated or saved later
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(HelpMessage = 'Free/busy status.')]
        [ValidateNotNullOrEmpty()]
        [ews_legacyfreebusystatus[]]$FreeBusyStatus = [System.Enum]::GetValues([ews_legacyfreebusystatus]),
        [bool]$IsAllDayEvent = $true,
        [bool]$IsReminderSet = $false,
        [datetime]$Start = (Get-Date),
        [datetime]$End = (Get-Date),
        [string]$Subject,
        [string]$Location,
        [string]$Body
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
    
    Write-Verbose "$($FunctionName): Attempting to create an appointment"
    if ($FreeBusyStatus.count -gt 1) {
        $FreeBusyStatus = [Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus]::Free
    }
    # Construct Appointment
    $appt = [ews_appt]($EWSService)
    # $cstzone = [System.TimeZoneInfo]::FindSystemTimeZoneById(($EWSService.TimeZone).StandardName)
    # $appt.StartTimeZone = $cstzone
    $appt.LegacyFreeBusyStatus = $FreeBusyStatus
    $appt.IsReminderSet = $IsReminderSet
    $appt.IsAllDayEvent = $IsAllDayEvent
    if ($IsAllDayEvent) {
        $StartDate = (Get-Date ($Start.ToShortDateString() + ' 9:00 AM') -Format 's') + '-600'
        $EndDate = (Get-Date ($Start.ToShortDateString() + ' 5:00 PM') -Format 's') + '-600'

        $appt.Start = [DateTime]::Parse($StartDate)
        $appt.End = [DateTime]::Parse($EndDate)

        #$appt.Start = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date ($Start.ToShortDateString())).ToUniversalTime(), $cstzone)
        #$appt.End = ($appt.Start).AddHours(24)
    }
    else {
        $appt.Start = $Start
        $appt.End = $End
    }
    $appt.Subject = $Subject
    $appt.Location = $Location
    $appt.Body = $Body

    return $appt
}