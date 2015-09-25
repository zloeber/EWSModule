function Get-EWSCalenderViewAppointments {
    <#
    .SYNOPSIS
        Uses a slower method for accessing and returning calendar appointments
    .DESCRIPTION
        Uses a slower method for accessing and returning calendar appointments
    .PARAMETER EWSService
        Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
        Mailbox to target. If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.
    .PARAMETER StartRange
        Start of when to look for appointments.
    .PARAMETER EndRange
        End of when to look for appointments.

    .EXAMPLE
        PS > 
        PS > 

        Description
        -----------
        TBD

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [string]$Mailbox = '',
        [datetime]$StartRange = (Get-Date),
        [datetime]$EndRange = ((Get-Date).AddMonths(12))
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
    
    try {
        $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
    }
    catch {
        throw "$($FunctionName): Unable to get targeted mailbox"
    }
    
    Write-Verbose "$($FunctionName): Attempting to gather calendar entries for $($email)"
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