#region Private Variables
# Use this variable for any path-sepecific actions (like loading dlls and such) to ensure it will work in testing and after being built
$MyModulePath = $(
    Function Get-ScriptPath {
        $Invocation = (Get-Variable MyInvocation -Scope 1).Value
        if($Invocation.PSScriptRoot) {
            $Invocation.PSScriptRoot
        }
        Elseif($Invocation.MyCommand.Path) {
            Split-Path $Invocation.MyCommand.Path
        }
        elseif ($Invocation.InvocationName.Length -eq 0) {
            (Get-Location).Path
        }
        else {
            $Invocation.InvocationName.Substring(0,$Invocation.InvocationName.LastIndexOf("\"));
        }
    }

    Get-ScriptPath
)

# Track if we have gone through the Initialize-EWS function yet.
[bool]$EWSModuleInitialized = $false

# Current script path
[string]$ScriptPath = Split-Path (get-variable myinvocation -scope script).value.Mycommand.Definition -Parent

# Used to track if we are setting SSL work around globally or not.
[bool]$IsSSLWorkAroundInPlace = $false

# A bunch of custom type accelerators to make the code look much less insane ( or more depending how you look at it I suppose )
$EWSAccels = @{
    'ews_basepropset'='Microsoft.Exchange.WebServices.Data.BasePropertySet'
    'ews_connidtype'='Microsoft.Exchange.WebServices.Data.ConnectingIdType'
    'ews_extendedpropset'='Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet'
    'ews_extendedpropdef'='Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition'
    'ews_propset'='Microsoft.Exchange.WebServices.Data.PropertySet'
    'ews_folder'='Microsoft.Exchange.WebServices.Data.Folder'
    'ews_calendarfolder'='Microsoft.Exchange.WebServices.Data.CalendarFolder'
    'ews_calendarview'='Microsoft.Exchange.WebServices.Data.CalendarView'
    'ews_folderid'='Microsoft.Exchange.WebServices.Data.FolderId'
    'ews_folderview'='Microsoft.Exchange.WebServices.Data.FolderView'
    'ews_impersonateuserid'='Microsoft.Exchange.WebServices.Data.ImpersonatedUserId'
    'ews_mailbox'='Microsoft.Exchange.WebServices.Data.Mailbox'
    'ews_mapiproptype'='Microsoft.Exchange.WebServices.Data.MapiPropertyType'
    'ews_operator'='Microsoft.Exchange.WebServices.Data.LogicalOperator'
    'ews_resolvenamelocation'='Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation'
    'ews_schema_appt'='Microsoft.Exchange.WebServices.Data.AppointmentSchema'
    'ews_schema_folder'='Microsoft.Exchange.WebServices.Data.FolderSchema'
    'ews_schema_item'='Microsoft.Exchange.WebServices.Data.ItemSchema'
    'ews_searchfilter'='Microsoft.Exchange.WebServices.Data.SearchFilter'
    'ews_searchfilter_collection'='Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection'
    'ews_searchfilter_isequalto'='Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo'
    'ews_searchfilter_isgreaterthanorequalto'='Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo'
    'ews_searchfilter_islessthanorequalto'='Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo'
    'ews_searchfilter_exists'='Microsoft.Exchange.WebServices.Data.SearchFilter+Exists'
    'ews_service'='Microsoft.Exchange.WebServices.Data.ExchangeService'
    'ews_webcredential'='Microsoft.Exchange.WebServices.Data.WebCredentials'
    'ews_wellknownfolder'='Microsoft.Exchange.WebServices.Data.WellKnownFolderName'
    'ews_itemview'='Microsoft.Exchange.WebServices.Data.ItemView'
    'ews_appttype'='Microsoft.Exchange.WebServices.Data.AppointmentType'
    'ews_appt'='Microsoft.Exchange.WebServices.Data.Appointment'
    'ews_deletemode'='Microsoft.Exchange.WebServices.Data.DeleteMode'
    'ews_sendcancellationmode'='Microsoft.Exchange.WebServices.Data.SendCancellationsMode'
    'ews_conflictresolutionmode'='Microsoft.Exchange.WebServices.Data.ConflictResolutionMode'
    'ews_sendinvitationorcancellationsmode'='Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode'
    'ews_legacyfreebusystatus'='Microsoft.Exchange.WebServices.Data.LegacyFreeBusyStatus'
    'ews_autod'='Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService'
    'ews_usersettingname'='Microsoft.Exchange.WebServices.Autodiscover.UserSettingName'
    'ews_contact' = 'Microsoft.Exchange.WebServices.Data.Contact'
    'ews_mailboxtype' = 'Microsoft.Exchange.WebServices.Data.MailboxType'
    'ews_exchver' = 'Microsoft.Exchange.WebServices.Data.ExchangeVersion'
}

$ewsdllpaths = @( 
    "$($MyModulePath)\Microsoft.Exchange.WebServices.dll",
    "$($ScriptPath)\Microsoft.Exchange.WebServices.dll",
    'C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll',
    'C:\Program Files\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll')

$modEWSService = $null
$modCertCallback = $null
#endregion Private Variables