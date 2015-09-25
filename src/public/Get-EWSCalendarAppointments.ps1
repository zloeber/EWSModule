function Get-EWSCalendarAppointments {
    # 
    <#
    .SYNOPSIS
        Uses the much faster FindItems as opposed to FindAppointments to return calendar appointments.
    .DESCRIPTION
        Uses the much faster FindItems as opposed to FindAppointments to return calendar appointments.
    .PARAMETER EWSService
        Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
        Mailbox to target. If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.
    .PARAMETER FolderPath
        Path of folder in the form of /folder1/folder2
    .PARAMETER StartsAfter
        Start date for the appointment(s) must be after this date
    .PARAMETER StartsBefore
        Start date for the appointment(s) must be before this date
    .PARAMETER EndsAfter
        nd date for the appointment(s) must be after this date
    .PARAMETER EndsBefore
        nd date for the appointment(s) must be before this date
    .PARAMETER CreatedAfter
        Only appointments created after the given date will be returned
    .PARAMETER CreatedBefore
        Only appointments created before the given date will be returned
    .PARAMETER LastOccurrenceAfter
        Only recurring appointments with a last occurrence date after the given date will be returned
    .PARAMETER LastOccurrenceBefore
        Only recurring appointments with a last occurrence date before the given date will be returned
    .PARAMETER IsRecurring
        If this switch is present, only recurring appointments are returned
    .PARAMETER ExtendedProperties
        Filter results by custom extended properties.
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
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [Parameter(HelpMessage="Mailbox to search - if omitted the EWS connection account ID is used (or impersonated account if set).")] 
        [string]$Mailbox = '',
        [Parameter(HelpMessage="Folder to search - if omitted, the mailbox calendar folder is assumed")] 
        [string]$FolderPath,
        [Parameter(HelpMessage="Subject of the appointment(s) being searched")] 
        [string]$Subject,
        [Parameter(HelpMessage="Start date for the appointment(s) must be after this date")] 
        [datetime]$StartsAfter,
        [Parameter(HelpMessage="Start date for the appointment(s) must be before this date")] 
        [datetime]$StartsBefore, 
        [Parameter(HelpMessage="End date for the appointment(s) must be after this date")] 
        [datetime]$EndsAfter, 
        [Parameter(HelpMessage="End date for the appointment(s) must be before this date")] 
        [datetime]$EndsBefore, 
        [Parameter(HelpMessage="Only appointments created before the given date will be returned")] 
        [datetime]$CreatedBefore, 
        [Parameter(HelpMessage="Only appointments created after the given date will be returned")] 
        [datetime]$CreatedAfter, 
        [Parameter(HelpMessage="Only recurring appointments with a last occurrence date before the given date will be returned")] 
        [datetime]$LastOccurrenceBefore, 
        [Parameter(HelpMessage="Only recurring appointments with a last occurrence date after the given date will be returned")] 
        [datetime]$LastOccurrenceAfter, 
        [Parameter(HelpMessage="If this switch is present, only recurring appointments are returned")]
        [switch]$IsRecurring,
        [Parameter(HelpMessage='Search for extended properties.')]
        [ews_extendedpropdef[]]$ExtendedProperties
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

    $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox

    Write-Verbose "Get-EWSCalendarEnties: Attempting to gather calendar entries for $($email)"

    $MailboxToAccess = new-object ews_mailbox($email)

    if ([string]::IsNullOrEmpty($FolderPath)) {
        $FolderID = new-object ews_folderid([ews_wellknownfolder]::Calendar, $MailboxToAccess)
    }

    $EWSCalFolder = [ews_calendarfolder]::Bind($EWSService, $FolderID)
    $view = New-Object ews_itemview(500, 0)
    
    $offset = 0 
    $moreItems = $true
    $filters = @()
    
	#region Build Extended Property Set for Item Results
	# Build the Item Property Set and then add the Properties that we want
	$customPropSet = New-Object -TypeName ews_propset([ews_basepropset]::FirstClassProperties)

	# Define the Item Extended Properties and add to collection (if defined)
    if ($ExtendedProperties -ne $null) {
        $ExtendedProperties | Foreach {
            $customPropSet.Add($_)
            $filters += New-Object ews_searchfilter_exists($_)
        }
    }
    $customPropSet.Add([ews_schema_item]::ID)
    $customPropSet.Add([ews_schema_item]::Subject)
    $customPropSet.Add([ews_schema_appt]::Start)
    $customPropSet.Add([ews_schema_appt]::End)
    $customPropSet.Add([ews_schema_item]::DateTimeCreated)
    $customPropSet.Add([ews_schema_appt]::AppointmentType)
    $view.PropertySet = $customPropSet
    #endregion Build Extended Property Set for Item Results

    # Set the search filter - this limits some of the results, not all the options can be filtered 
    if ($createdBefore -ne $Null) { 
        $filters += New-Object ews_searchfilter_IsLessThanOrEqualTo([ews_schema_item]::DateTimeCreated, $CreatedBefore) 
    }
    if (-not [string]::IsNullOrEmpty($Subject)) { 
        $filters += New-Object ews_searchfilter_isequalto([ews_schema_item]::Subject, $Subject) 
    }
    if ($createdAfter -ne $Null) { 
        $filters += New-Object ews_searchfilter_IsGreaterThanOrEqualTo([ews_schema_item]::DateTimeCreated, $createdBefore) 
    } 
    if ($startsBefore -ne $Null) { 
        $filters += New-Object ews_searchfilter_IsLessThanOrEqualTo([ews_schema_appt]::Start, $startsBefore) 
    } 
    if ($startsAfter -ne $Null) { 
        $filters += New-Object ews_searchfilter_IsGreaterThanOrEqualTo([ews_schema_appt]::Start, $startsAfter) 
    } 
    if ($endsBefore -ne $Null) { 
        $filters += New-Object ews_searchfilter_IsLessThanOrEqualTo([ews_schema_appt]::End, $endsBefore) 
    } 
    if ($endsAfter -ne $Null) { 
        $filters += New-Object ews_searchfilter_IsGreaterThanOrEqualTo([ews_schema_appt]::End, $endsAfter) 
    }
    if ($IsRecurring) {
        $filters += New-Object ews_searchfilter_isequalto([ews_schema_appt]::IsRecurring,$true)
    }
    $searchFilter = $Null
    if ( $filters.Count -gt 0 ) { 
        $searchFilter = New-Object ews_searchfilter_collection([ews_operator]::And) 
        foreach ($filter in $filters) {
            $searchFilter.Add($filter) 
        } 
    } 
 
    # Now retrieve the matching items and process 
    while ($moreItems) { 
        # Get the next batch of items to process 
        if ( $searchFilter ) { 
            $results = $EWSCalFolder.FindItems($searchFilter, $view) 
        } 
        else { 
            $results = $EWSCalFolder.FindItems($view) 
        } 
        $moreItems = $results.MoreAvailable 
        $view.Offset = $results.NextPageOffset 

        $results
    }
}