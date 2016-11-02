function New-EWSCalendarEntry {
    <# 
    .SYNOPSIS 
    Creates an appointment object that can be manipulated or saved.
    .DESCRIPTION 
    Creates an appointment object that can be manipulated or saved.
    .PARAMETER EWSService
    Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER FreeBusyStatus
    FreeBusy status for the appointment. Can be 'Free','Tentative','Busy','OOF','WorkingElsewhere', or 'NoData'. Defaults to 'Free'.
    .PARAMETER IsAllDayEvent
    Set the flag to mark the appointment as an all day event.
    .PARAMETER IsReminderSet
    Set the flag to mark the appointment to have a default reminder.
    .PARAMETER Start
    Start time of the appointment.
    .PARAMETER End
    End time of the appointment.
    .PARAMETER Subject
    Appointment subject line.
    .PARAMETER Location
    Appointment location.
    .PARAMETER Body
    Body of the appointment.
    .EXAMPLE
    $CalEntry = New-EWSCalendarEntry -IsAllDayEvent $true -Subject 'My Event' -Location 'Elsewhere'

    Creates a new calendar entry as an all day event called 'My Event' and stores it in $CalEntry

    .LINK
    http://www.the-little-things.net/

    .LINK
    https://www.github.com/zloeber/EWSModule

    .NOTES
    Author: Zachary Loeber
    Requires: Powershell 3.0
    Version History
    1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage = 'Free/busy status.')]
        [ValidateSet('Free','Tentative','Busy','OOF','WorkingElsewhere','NoData')]
        [ews_legacyfreebusystatus]$FreeBusyStatus = [ews_legacyfreebusystatus]::Free,
        [parameter(Position=2)]
        [bool]$IsAllDayEvent = $false,
        [parameter(Position=3)]
        [bool]$IsReminderSet = $false,
        [parameter(Position=4)]
        [datetime]$Start = (Get-Date),
        [parameter(Position=5)]
        [datetime]$End = (Get-Date),
        [parameter(Position=6)]
        [string]$Subject,
        [parameter(Position=7)]
        [string]$Location,
        [parameter(Position=8)]
        [string]$Body
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
    
    Write-Verbose "$($FunctionName): Attempting to create an appointment"
    if ($FreeBusyStatus.count -gt 1) {
        $FreeBusyStatus = [ews_legacyfreebusystatus]::Free
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