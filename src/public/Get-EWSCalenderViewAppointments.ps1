function Get-EWSCalenderViewAppointments {
    # uses a slower method for accessing appointments
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService = (Get-EWSService),
        [string]$Mailbox = '',
        [datetime]$StartRange = (Get-Date),
        [datetime]$EndRange = ((Get-Date).AddMonths(12))
    )
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw 'EWS Module has not been initialized. Try running Initialize-EWS to rectify.'
    }
    
    $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
    
    Write-Verbose "Get-EWSCalendarEnties: Attempting to gather calendar entries for $($email)"
    $MailboxToAccess = new-object ews_mailbox($email)

    $FolderID = new-object ews_folderid([ews_wellknownfolder]::Calendar, $MailboxToAccess)

    $EWSCalFolder = [ews_calendarfolder]::Bind($EWSService, $FolderID)
    $propsetfc = [ews_basepropset]::FirstClassProperties
    $Calview = new-object ews_calendarview($StartRange, $EndRange, 1000)
    $Calview.PropertySet = $propsetfc

    $appointments = @()
    $CalSearchResult = $EWSService.FindAppointments($EWSCalFolder.id, $Calview)
    $appointments += $CalSearchResult

    while($CalSearchResult.MoreAvailable) {
        $calview.StartDate = $CalSearchResult.Items[$CalSearchResult.Items.Count-1].Start
        $CalSearchResult = $EWSService.FindAppointments($EWSCalFolder.id, $Calview)
        $appointments += $CalSearchResult
    }

    $appointments.GetEnumerator()
}