#region Private Variables
# Privately set variable to track if we have gone through the Initialize-EWS function yet.
[bool]$EWSModuleInitialized = $false

# Current script path
[string]$ScriptPath = Split-Path (get-variable myinvocation -scope script).value.Mycommand.Definition -Parent

# Used to track if we are setting SSL work around globally or not.
[bool]$IsSSLWorkAroundInPlace = $false

# A bunch of custom type accelerators to make the code look much less insane
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
}

$ewsdllpaths = @( 
    "$($ScriptPath)\Microsoft.Exchange.WebServices.dll",
    'C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll',
    'C:\Program Files\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll')

$modEWSService = $null
$modCertCallback = $null
#endregion Private Variables

#region Dependancies
Get-ChildItem $ScriptPath/src -Recurse -Filter "*.ps1" -File | Foreach { 
    Write-Verbose "Dot sourcing file: $($_.Name)"
    . $_.FullName
}
#endregion Depenencies

#region Methods

function Import-EWSDll {
    <#
    .SYNOPSIS
    Load EWS dlls.

    .DESCRIPTION
    Load EWS dlls.
     
    .PARAMETER EWSManagedApiPath
    Full path to Microsoft.Exchange.WebServices.dll. If not provided we will try to load it from several best guess locations.
    
    .EXAMPLE
    Import-EWSDll
         
    .NOTES
    This function requires Exchange Web Services Managed API. From what I can tell you don't even need to install the msi. AS long
    as the Microsoft.Exchange.WebServices.dll file is extracted and available that should work.
    
    The EWS Managed API can be obtained from: http://www.microsoft.com/en-us/download/details.aspx?id=28952    
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=0)]
        [string]$EWSManagedApiPath
    )
    $FunctionName = $MyInvocation.MyCommand
    $ewspaths = @()
    if (-not (Get-EWSDllLoadState)) {
        if (-not [string]::IsNullOrEmpty($EWSManagedApiPath)) {
            $ewspaths += @($EWSManagedApiPath)
        }
        $ewspaths += $script:ewsdllpaths

        $EWSLoaded = $false
        foreach ($ewspath in $ewspaths) {
            try {
                if (-not $EWSLoaded) {
                    if (Test-Path $ewspath) {
                        Write-Verbose "$($FunctionName): Attempting to load $ewspath"
                        Import-Module -Name $ewspath -ErrorAction:Stop -Global
                        $EWSLoaded = $true
                    }
                }
            }
            catch {}
        }
    }
    else {
        Write-Verbose ("$($FunctionName): EWS dll already Loaded!")
    }
}

function Get-EWSDllLoadState {
    if (-not (get-module Microsoft.Exchange.WebServices)) {
        return $false
    }
    else {
        return $true
    }
}

function Get-EWSModuleInitializationState {
    return $script:EWSModuleInitialized
}

function Get-ServerCertificateValidationCallback {
    return $script:modCertCallback
}

function Set-ServerCertificateValidationCallback ($CertCallback) {
    $script:modCertCallback = $CertCallback
}

function Get-EWSService {
    return $script:modEWSService
}

function Set-EWSService ([ews_service]$ConnectedService) {
    $script:modEWSService = $ConnectedService
}

function Set-EWSModuleInitializationState ([bool]$State) {
    $script:EWSModuleInitialized = $State
}

function Initialize-EWS {
    <#
    .SYNOPSIS
    Load EWS dlls and create type accelerators for other functions.

    .DESCRIPTION
    Load EWS dlls and create type accelerators for other functions.
     
    .PARAMETER EWSManagedApiPath
    Full path to Microsoft.Exchange.WebServices.dll. If not provided we will try to load it from several best guess locations.
    
    .PARAMETER Uninitialize
    Remove previously added type-accelerators.
     
    .EXAMPLE
    Initialize-EWS
         
    .NOTES
    This function requires Exchange Web Services Managed API. From what I can tell you don't even need to install the msi. AS long
    as the Microsoft.Exchange.WebServices.dll file is extracted and available that should work.
    
    The EWS Managed API can be obtained from: http://www.microsoft.com/en-us/download/details.aspx?id=28952    
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=0)]
        [string]$EWSManagedApiPath,
        [parameter(Position=1)]
        [switch]$Uninitialize
    )
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not $Uninitialize) {
        Import-EWSDll -EWSManagedApiPath $EWSManagedApiPath
        if (Get-EWSDllLoadState) {
            if (-not (Get-EWSModuleInitializationState)) {
                # Setup a bunch of type accelerators to make this mess easier to understand (slightly)
                $accelerators = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')
                
                Add-Type -AssemblyName Microsoft.Exchange.WebServices
                Write-Verbose ("$($FunctionName): Attempting to create type accelerators.")
                foreach ($Key in ($script:EWSAccels).Keys) {
                    Write-Verbose "$($FunctionName): Adding type accelerator - $Key for the type $($Script:EWSAccels[$Key])"
                    $accelerators::Add($Key,$script:EWSAccels[$Key])
                }
                
                # Powershell 5.0 needs this or nothing will work (dammit!)
                if ($PSVersionTable.PSVersion.Major -eq 5) {
                    $builtinfield = $accelerators.GetField('builtinTypeAccelerators',[System.Reflection.BindingFlags]'Static,NonPublic')
                    $builtinfield.SetValue($builtinfield,$accelerators::Get)
                }

                Set-EWSModuleInitializationState $true
                return $true
            }
            else {
                return $true
            }
        }
        else {
            throw "$($FunctionName): Cant load EWS module. Please verify it is installed or manually provide the path to Microsoft.Exchange.WebServices.dll"
        }
    }
    else {
        # Uninitialize EWS
        if (Get-EWSModuleInitializationState) {
            $accelerators = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')
            $accelkeyscopy = @{}
            $accelkeys.Keys | Where {$_ -like 'ews_*'} | Foreach { $accelkeyscopy.$_ = $accelkeys[$_] }
            foreach ( $key in $accelkeyscopy.Keys ) {
                Write-Verbose "UnInitialize-EWS: Removing type accelerator - $($key)"
                $accelerators::Remove("$($key)") | Out-Null
            }
            Write-Verbose ("$($FunctionName): Custom type accelerators removed!")
            Set-EWSModuleInitializationState $false
        }
        if (Get-EWSDllLoadState) {
            Remove-Module Microsoft.Exchange.WebServices
            Write-Verbose ("$($FunctionName): EWS dll Unloaded!")
        }

        return $true
    }
}

function Connect-EWS {
    [CmdLetBinding(DefaultParameterSetName='Default')]
    param(
        [parameter(Mandatory=$True,ParameterSetName='CredentialString', HelpMessage='Alternate credential username.')]
        [string]$UserName,
        [parameter(Mandatory=$True,ParameterSetName='CredentialString')]
        [string]$Password,
        [parameter(ParameterSetName='CredentialString')]
        [string]$Domain,
        [parameter(Mandatory=$True,ParameterSetName='CredentialObject')]
        [alias('Creds')]
        [System.Management.Automation.PSCredential]$Credential,
        [parameter(ParameterSetName='CredentialString')]
        [parameter(ParameterSetName='CredentialObject')]
        [parameter(ParameterSetName='Default')]
        [ValidateSet('Exchange2013_SP1','Exchange2013','Exchange2010_SP2','Exchange2010_SP1','Exchange2010','Exchange2007_SP1')]
        [string]$ExchangeVersion = 'Exchange2010_SP2',
        [parameter(ParameterSetName='CredentialString')]
        [parameter(ParameterSetName='CredentialObject', HelpMessage='Use statically set ews url. Autodiscover is attempted otherwise.')]
        [parameter(ParameterSetName='Default')]
        [string]$EwsUrl='',
        [parameter(ParameterSetName='CredentialString')]
        [parameter(ParameterSetName='CredentialObject')]
        [parameter(ParameterSetName='Default')]
        [switch]$EWSTracing,
        [parameter(ParameterSetName='CredentialString')]
        [parameter(ParameterSetName='CredentialObject')]
        [parameter(ParameterSetName='Default')]
        [switch]$IgnoreSSLCertificate
    )
    $FunctionName = $MyInvocation.MyCommand
    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }

    #Load credential info
    switch ($PSCmdlet.ParameterSetName) {
        'CredentialObject' {
            $UserName= $Credential.GetNetworkCredential().UserName
            $Password = $Credential.GetNetworkCredential().Password
            $Domain = $Credential.GetNetworkCredential().Domain
        }
    }

    if ($IgnoreSSLCertificate -and (-not $script:IsSSLWorkAroundInPlace)) {
        Write-Verbose "$($FunctionName): Ignoring any SSL certificate errors"
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    }

    try {        
        Write-Verbose "$($FunctionName): Creating EWS Service object with exchange version of $ExchangeVersion"
        $enumExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion
        $tempEWSService = new-object ews_service($enumExchVer) -ErrorAction Stop
    }
    catch {
        Write-Error ('Connect-EWS: Cannot create EWS Service with the following defined Exchange version- {0}' -f $ExchangeVersion)
        throw ("$($FunctionName): Full Error - $($_.Exception.Message)")
    }
    
    # If an alternate credential has been passed setup accordingly
    if ($UserName) {
        if (-not [string]::IsNullOrEmpty($Domain)) {
            #If a domain is presented then use that as well
            $tempEWSService.Credentials = New-Object ews_webcredential($UserName,$Password,$Domain) -ErrorAction Stop
        }
        else {
            #Otherwise leave the domain blank
            $tempEWSService.Credentials = New-Object ews_webcredential($UserName,$Password) -ErrorAction Stop
        }
    }

    # Otherwise try to use the current account
    else {
        $tempEWSService.UseDefaultCredentials = $true
    }
    
    if ($EWSTracing) {
        Write-Verbose "$($FunctionName): EWS Tracing enabled"
        $tempEWSService.traceenabled = $true
    }

    # If an ews url was defined then use that first
    if (-not [string]::IsNullOrEmpty($EwsUrl)) {
        Write-Verbose "$($FunctionName): Using the specifed EWS URL of $EwsUrl"
        $tempEWSService.URL = New-Object Uri($EwsUrl) -ErrorAction Stop
    }
    # Otherwise try to use autodiscover to get the url
    else {
        $AutoDiscoverSplat = @{}
        if ($UserName) {
            # If using an alternate userid then try autodiscover with it, otherwise the current account is used
            $AutoDiscoverSplat.UserID = $UserName
        }
        try {
            $AutodiscoverAccount = Get-EmailAddressFromAD @AutoDiscoverSplat
        }
        catch {
            throw "$($FunctionName): Unable to find a primary smtp account with this account in AD. Try using the email format for the user login ID instead."
        }
        try {
            Write-Verbose "$($FunctionName): Performing autodiscover for - $AutodiscoverAccount"
            $AutodiscoverInfo = Test-EWSAutodiscover -EmailAddress $AutodiscoverAccount -Credential $Credential
            $tempEWSService.URL = New-Object Uri($AutodiscoverInfo.ExternalEwsUrl) -ErrorAction Stop
        }
        catch {
            throw "$($FunctionName): EWS Url not specified and autodiscover failed, bummer."
        }
    }
    
    Set-EWSService $tempEWSService
    $tempCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
    if ($tempCallback -ne $null) {
        Write-Verbose "$($FunctionName): ServerCertificateValidationCallback being set to null for this session."
        Set-ServerCertificateValidationCallback [System.Net.ServicePointManager]::ServerCertificateValidationCallback
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
    }
    
    return $true
}

function Get-EmailAddressFromAD {
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='ID to lookup. Defaults to current users SID')]
        [string]$UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    )
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Test-EmailAddressFormat $UserID)) {        
        try {
            if (Test-UserSIDFormat $UserID) {
                $user = [ADSI]"LDAP://<SID=$sid>"
                $retval = $user.Properties.mail
            }
            else {
                $strFilter = "(&(objectCategory=User)(samAccountName=$($UserID)))"
                $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
                $objSearcher.Filter = $strFilter
                $objPath = $objSearcher.FindOne()
                $objUser = $objPath.GetDirectoryEntry()
                $retval = $objUser.mail
            }
        }
        catch {
            Write-Debug ("$($FunctionName): Full Error - $($_.Exception.Message)")
            throw "$($FunctionName): Cannot get directory information for $UserID"
        }
        if ([string]::IsNullOrEmpty($retval)) {
            Write-Verbose "$($FunctionName): Cannot determine the primary email address for - $UserID"
            throw "$($FunctionName): Autodiscover failure - No email address associated with current user."
        }
        else {
            return $retval
        }
    }
    else {
        return $UserID
    }
}

function Get-EWSCalendarAppointments {
    # Use FindItems as opposed to FindAppointments
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService = (Get-EWSService),
        [Parameter(HelpMessage="Mailbox to search - if omitted the EWS connection account ID is used (or impersonated account if set).")] 
        [string]$Mailbox = '',
        [Parameter(HelpMessage="Folder to search - if omitted, the mailbox calendar folder is assumed")] 
        $FolderPath,
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
        [Parameter(HelpMessage='Search for extended properties being set.')]
        [ews_extendedpropdef[]]$ExtendedProperties
    )
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw 'EWS Module has not been initialized. Try running Initialize-EWS to rectify.'
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

function Get-EWSFolder {
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(Position=1, HelpMessage='Mailbox of folder.')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Folder path.')]
        [string]$FolderPath,
        [parameter(Position=2, HelpMessage='Public Folder Path?')]
        [switch]$PublicFolder
    )
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    
    if ($EWSService -eq $null) {
        Write-Verbose "$($FunctionName): Using module local ews service object"
        $EWSService = Get-EWSService
        Write-Verbose "$($FunctionName): URL targeted = $($EWSService.URL)"
    }
    
    if ($EWSService -eq $null) {
        throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
    }
    
    # Return a reference to a folder specified by path 
    if ($PublicFolders) { 
        $mbx = '' 
        $Folder = [ews_folder]::Bind($EWSService, [ews_wellknownfolder]::PublicFoldersRoot) 
    } 
    else {
        $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
        $mbx = New-Object ews_mailbox( $email ) 
        $folderId = New-Object ews_folderid([ews_wellknownfolder]::MsgFolderRoot, $mbx ) 
        $Folder = [ews_folder]::Bind($EWSService, $folderId) 
    } 
 
    if ($FolderPath -ne '\') {
        $PathElements = $FolderPath -split '\\' 
        For ($i=0; $i -lt $PathElements.Count; $i++) { 
            if ($PathElements[$i]) { 
                $View = New-Object  ews_folderview(2,0) 
                $View.PropertySet = [ews_basepropset]::IdOnly
                $SearchFilter = New-Object ews_searchfilter_isequalto([ews_schema_folder]::DisplayName, $PathElements[$i])
                $FolderResults = $Folder.FindFolders($SearchFilter, $View) 
                if ($FolderResults.TotalCount -ne 1) { 
                    # We have either none or more than one folder returned... Either way, we can't continue 
                    $Folder = $null 
                    Write-Verbose "$($FunctionName): Failed to find $($PathElements[$i]), path requested was $FolderPath"
                    break 
                }
                 
                if (-not [String]::IsNullOrEmpty(($mbx))) {
                    $folderId = New-Object ews_folderid($FolderResults.Folders[0].Id, $mbx ) 
                    $Folder = [ews_folder]::Bind($service, $folderId) 
                } 
                else {
                    $Folder = [ews_folder]::Bind($service, $FolderResults.Folders[0].Id) 
                } 
            } 
        } 
    }

    return $Folder 
}

function Get-EWSTargettedMailbox {
    # Supplemental function 
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(Position=1, HelpMessage='Mailbox you are targeting.')]
        [string]$Mailbox
    )
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    
    if ($EWSService -eq $null) {
        Write-Verbose "$($FunctionName): Using module local ews service object"
        $EWSService = Get-EWSService
        Write-Verbose "$($FunctionName): URL targeted = $($EWSService.URL)"
    }
    
    if ($EWSService -eq $null) {
        throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
    }
    
    if (-not [string]::IsNullOrEmpty($Mailbox)) {
        if (Test-EmailAddressFormat $Mailbox) {
            $email = $Mailbox
        }
        else {
            try {
                $email = Get-EmailAddressFromAD $Mailbox
            }
            catch {
                throw "$($FunctionName): Unable to get a mailbox"
            }
        }
    }
    else {
        if ($EWSService.ImpersonatedUserId -ne $null) {
            $impID = $EWSService.ImpersonatedUserId.Id
        }
        else {
            $impID = $EWSService.Credentials.Credentials.UserName
        }
        
        if (-not (Test-EmailAddressFormat $impID)) {
            try {
                $email = ($EWSService.ResolveName("smtp:$($ImpID)@",[ews_resolvenamelocation]::DirectoryOnly, $false)).Mailbox -creplace '(?s)^.*\:', '' -creplace '>',''
            }
            catch {
                throw "$($FunctionName): Unable to find a mailbox with this account."
            }
        }
        else {
            $email = $impID
        }
    }
    
    return $email
}

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

function Set-EWSMailboxImpersonation {
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(Position=1, Mandatory=$True, HelpMessage='Mailbox to impersonate.')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Do not attempt to validate rights against this mailbox (can speed up operations)')]
        [switch]$SkipValidation
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

    if (Test-EmailAddressFormat $Mailbox) {
        $enumType = [ews_connidtype]::SmtpAddress
    }
    else {
        $enumType = [ews_connidtype]::PrincipalName
    }
    try {
        $EWSService.ImpersonatedUserId = New-Object ews_impersonateuserid($enumType,$Mailbox)
        if (-not $SkipValidation) {
            $InboxFolder= new-object ews_folderid([ews_wellknownfolder]::Inbox,$Mailbox)
            $Inbox = [ews_folder]::Bind($EWSService,$InboxFolder)
        }
    }
    catch {
        Write-Error ('Set-EWSMailboxImpersonation: Unable to impersonate {0}, check to see that you have adequately assigned permissions to impersonate this account.' -f $Mailbox)
        throw $_.Exception.Message  
    }
}

function Test-EWSAutodiscover {
    <#
    .SYNOPSIS
        This function uses the EWS Managed API to test the Exchange Autodiscover service.

    .DESCRIPTION
        This function will retreive the Client Access Server URLs for a specified email address
        by querying the autodiscover service of the Exchange server.

    .PARAMETER  EmailAddress
        Specifies the email address for the mailbox that should be tested.

    .PARAMETER  Location
        Set to External by default, but can also be set to Internal. This parameter controls whether
        the internal or external URLs are returned.
        
    .PARAMETER  Credential
        Specifies a user account that has permission to perform this action. Type a user name, such as 
        "User01" or "Domain01\User01", or enter a PSCredential object, such as one from the Get-Credential cmdlet.
        
    .PARAMETER  TraceEnabled
        Use this switch parameter to enable tracing. This is used for debugging the XML response from the server.    
        
    .PARAMETER  IgnoreSsl
        Set to $true by default. If you do not want to ignore SSL warnings or errors, set this parameter to $false.
        
    .PARAMETER  Url
        You can use this parameter to manually specifiy the autodiscover url.        

    .EXAMPLE
        PS C:\> Test-Autodiscover -EmailAddress administrator@uclabs.ms -Location internal
        
        This example shows how to retrieve the internal autodiscover settings for a user.

    .EXAMPLE
        PS C:\> Test-Autodiscover -EmailAddress administrator@uclabs.ms -Credential $cred
        
        This example shows how to retrieve the external autodiscover settings for a user. You can
        provide credentials if you do not want to use the Windows credentials of the user calling
        the function.

    .LINK
        http://msdn.microsoft.com/en-us/library/dd633699%28v=EXCHG.80%29.aspx

    #>

    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true)]
        [String]$EmailAddress,

        [Parameter(Position=1)]
        [ValidateSet("Internal", "External")]
        [String]$Location = "External",      

        [Parameter(Position=2)]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter(Position=3)]
        [switch]$TraceEnabled,

        [Parameter(Position=4)]
        [String]$Url
    )

    if (-not (Get-EWSModuleInitializationState)) {
        throw 'EWS Module has not been initialized. Try running Initialize-EWS to rectify.'
    }
    
    $autod = New-Object ews_autod
    $autod.RedirectionUrlValidationCallback = {$true}
    $autod.TraceEnabled = $TraceEnabled

    if ($Credential) {
        $autod.Credentials = New-Object ews_webcredential -ArgumentList $Credential.UserName, $Credential.GetNetworkCredential().Password
    }

    if($Url) {
        $autod.Url = $Url
    }

    switch($Location) {
      'Internal' {
        $autod.EnableScpLookup = $true
        $response = $autod.GetUserSettings(
            $EmailAddress,
            [ews_usersettingname]::InternalRpcClientServer,
            [ews_usersettingname]::InternalEcpUrl,
            [ews_usersettingname]::InternalEwsUrl,
            [ews_usersettingname]::InternalOABUrl,
            [ews_usersettingname]::InternalUMUrl,
            [ews_usersettingname]::InternalWebClientUrls
        )
        
        New-Object PSObject -Property @{
            RpcClientServer = $response.Settings[[ews_usersettingname]::InternalRpcClientServer]
            InternalOwaUrl = $response.Settings[[ews_usersettingname]::InternalWebClientUrls].urls[0].url
            InternalEcpUrl = $response.Settings[[ews_usersettingname]::InternalEcpUrl]
            InternalEwsUrl = $response.Settings[[ews_usersettingname]::InternalEwsUrl]
            InternalOABUrl = $response.Settings[[ews_usersettingname]::InternalOABUrl]
            InternalUMUrl = $response.Settings[[ews_usersettingname]::InternalUMUrl]
        }
      }
      'External' {
        $autod.EnableScpLookup = $false
        $response = $autod.GetUserSettings(
            $EmailAddress,
            [ews_usersettingname]::ExternalMailboxServer,
            [ews_usersettingname]::ExternalEcpUrl,
            [ews_usersettingname]::ExternalEwsUrl,
            [ews_usersettingname]::ExternalOABUrl,
            [ews_usersettingname]::ExternalUMUrl,
            [ews_usersettingname]::ExternalWebClientUrls
        )
        
        New-Object PSObject -Property @{
            HttpServer = $response.Settings[[ews_usersettingname]::ExternalMailboxServer]
            ExternalOwaUrl = $response.Settings[[ews_usersettingname]::ExternalWebClientUrls].urls[0].url
            ExternalEcpUrl = $response.Settings[[ews_usersettingname]::ExternalEcpUrl]
            ExternalEwsUrl = $response.Settings[[ews_usersettingname]::ExternalEwsUrl]
            ExternalOABUrl = $response.Settings[[ews_usersettingname]::ExternalOABUrl]
            ExternalUMUrl = $response.Settings[[ews_usersettingname]::ExternalUMUrl]
        }
      }
    }
}

function Install-EWSDll {
    <#
    .SYNOPSIS
    Attempts to download and extract the ews dll needed for this library.

    .DESCRIPTION
    Attempts to download and extract the ews dll needed for this library.
    #>
    [CmdLetBinding()]
    param(
        [string]$source = 'http://download.microsoft.com/download/8/9/9/899EEF2C-55ED-4C66-9613-EE808FCF861C/EwsManagedApi.msi',
        [string]$destination,
        [switch]$SkipDownload
    )
    
    if ([string]::IsNullOrEmpty($destination) -or (-not (Test-Path $destination))) {
        $destination = (Split-Path (get-variable myinvocation -scope script).value.Mycommand.Definition -Parent) + "\EwsManagedApi.MSI"
    }
    if (-not $SkipDownload) {
        try {
            $splatparam = @{
                'source' = $source
                'destination' = $destination
            }
            Write-Verbose 'Install-EWSDll: Attempting to download MSI.'
            $Download = Get-WebFile @splatparam
        }
        catch {
            throw 'Install-EWSDll: Unable to download the EWS install file!'
        }
        $DestPath = (Split-Path $Download -Parent) + "\EWSFiles\"
    }
    else {
        $DestPath = ".\EWSFiles\"
    }
    Write-Verbose "Install-EWSDll: Attempting to extract MSI to $($DestPath)."
    Invoke-MSIExec /quiet /a $Download /qn TARGETDIR=$DestPath

    if (Test-Path ("$DestPath\Microsoft.Exchange.WebServices.dll")) {
        Write-Verbose "Install-EWSDll: Copying MSI back to the download path."
        Copy-Item -Path "$DestPath\Microsoft.Exchange.WebServices.dll" -Destination (Split-Path $Download -Parent)
        Remove-Item -Recurse -Force -Path $DestPath
        Remove-Item -Force $Download
    }
    else {
        throw 'Install-EWSDll: Unable to extract the EWS install file!'
    }
}

Function Set-EWSSSLIgnoreWorkaround {
    if (-not $script:IsSSLWorkAroundInPlace) {
        $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
        $Compiler=$Provider.CreateCompiler()
        $Params=New-Object System.CodeDom.Compiler.CompilerParameters
        $Params.GenerateExecutable=$False
        $Params.GenerateInMemory=$True
        $Params.IncludeDebugInformation=$False
        $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

        $TASource=@'
          namespace Local.ToolkitExtensions.Net.CertificatePolicy{
            public class TrustAll : System.Net.ICertificatePolicy {
              public TrustAll() { 
              }
              public bool CheckValidationResult(System.Net.ServicePoint sp,
                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                System.Net.WebRequest req, int problem) {
                return true;
              }
            }
          }
'@ 
        $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
        $TAAssembly=$TAResults.CompiledAssembly

        ## We now create an instance of the TrustAll and attach it to the ServicePointManager
        $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
        [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

        $script:IsSSLWorkAroundInPlace = $true
    }
}

function Get-EWSOofSettings {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias('Identity')]
        [String]$Mailbox
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
    try {
        $TargetedMailbox = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
    }
    catch {
        throw "$($FunctionName): Unable to target $Mailbox"
    }
    try {
        $oof = $EWSService.GetUserOofSettings($TargetedMailbox)
        New-Object PSObject -Property @{
            State = $oof.State
            ExternalAudience = $oof.ExternalAudience
            StartTime = $oof.Duration.StartTime
            EndTime = $oof.Duration.EndTime
            InternalReply = $oof.InternalReply
            ExternalReply = $oof.ExternalReply
            AllowExternalOof = $oof.AllowExternalOof
            Identity = $TargetedMailbox
        }
    }
    catch {
        throw "$($FunctionName): Unable to get out of office info for $TargetedMailbox"
    }
}

function Set-EWSOofSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [String]$Identity,
        [Parameter(Position=1)]
        [String]$State,
        [Parameter(Position=2)]
        [String]$ExternalAudience,
        [Parameter(Position=3)]
        [DateTime]$StartTime,
        [Parameter(Position=4)]
        [DateTime]$EndTime,        
        [Parameter(Position=5)]
        [String]$InternalReply,
        [Parameter(Position=6)]
        [String]$ExternalReply
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

    try {
        $TargetedMailbox = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
    }
    catch {
        throw "$($FunctionName): Unable to target $Mailbox"
    }
    try {
        $oof = $EWSService.GetUserOofSettings($TargetedMailbox)
    }
    catch {
        throw "$($FunctionName): Unable to get oof settings from $TargetMailbox"
    }

    if($StartTime -and $EndTime) {
        $Duration = New-Object Microsoft.Exchange.WebServices.Data.TimeWindow -arg $StartTime,$EndTime
        $PSBoundParameters.Duration = $Duration
        $PSBoundParameters.State = "Scheduled"
        [Void]$PSBoundParameters.remove("StartTime")
        [Void]$PSBoundParameters.remove("EndTime")
    }
    
    foreach($p in $PSBoundParameters.GetEnumerator()) {
        if($p.key -ne "Identity") {
            $oof."$($p.key)" = $p.value                
        }
    }
    $EWSService.SetUserOofSettings($Identity,$oof)
}
#endregion Methods

#region Module Export
Export-ModuleMember Install-EWSDll, Import-EWSDll, Initialize-EWS
Export-ModuleMember Connect-EWS, Set-EWSMailboxImpersonation, Get-EWSService, Set-EWSService
Export-ModuleMember Test-EWSAutodiscover, Get-EmailAddressFromAD
Export-ModuleMember Get-EWSDllLoadState, Get-EWSModuleInitializationState
Export-ModuleMember Get-EWSCalendarAppointments, Get-EWSCalenderViewAppointments
Export-ModuleMember Get-EWSFolder, Get-EWSTargetedFolder, Get-EWSTargettedMailbox
Export-ModuleMember New-EWSCalendarEntry, New-EWSExtendedProperty
Export-ModuleMember Set-EWSSSLIgnoreWorkaround
Export-ModuleMember Get-EWSOofSettings, Set-EWSOofSettings
#endregion Module Export

#region Module Cleanup
$ExecutionContext.SessionState.Module.OnRemove = {
    if ( Initialize-EWS -Uninitialize ) {}
    else { Write-Warning "Unable to uninitialize module" }
}
#endregion Module Cleanup