## Pre-Loaded Module code ##

#region Private Variables
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
    "$($ScriptPath)\Microsoft.Exchange.WebServices.dll",
    'C:\Program Files (x86)\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll',
    'C:\Program Files\Microsoft\Exchange\Web Services\2.1\Microsoft.Exchange.WebServices.dll')

$modEWSService = $null
$modCertCallback = $null
#endregion Private Variables



## PRIVATE MODULE FUNCTIONS AND DATA ##

function Convert-ByteArrayToString {
    <#
    .Synopsis
        Returns the string representation of a System.Byte[] array. ASCII string is the default, but Unicode, UTF7, UTF8 and UTF32 are available too.
    .Parameter ByteArray
        System.Byte[] array of bytes to put into the file. If you pipe this array in, you must pipe the [Ref] to the array. 
        Also accepts a single Byte object instead of Byte[].
    .Parameter Encoding
        Encoding of the string: ASCII, Unicode, UTF7, UTF8 or UTF32. ASCII is the default.
    .Link
        http://www.sans.org/windows-security/2010/02/11/powershell-byte-array-hex-convert
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [System.Byte[]]$ByteArray,
        [Parameter()]
        [string]$Encoding = 'ASCII'
    )
     
    switch ( $Encoding.ToUpper() ) {
    	 "ASCII"   { $EncodingType = "System.Text.ASCIIEncoding" }
    	 "UNICODE" { $EncodingType = "System.Text.UnicodeEncoding" }
    	 "UTF7"    { $EncodingType = "System.Text.UTF7Encoding" }
    	 "UTF8"    { $EncodingType = "System.Text.UTF8Encoding" }
    	 "UTF32"   { $EncodingType = "System.Text.UTF32Encoding" }
    	 Default   { $EncodingType = "System.Text.ASCIIEncoding" }
    }
    $Encode = new-object $EncodingType
    $Encode.GetString($ByteArray)
}
 



function Convert-HexStringToByteArray {
    <#
    .Synopsis
        Convert a string of hex data into a System.Byte[] array. An
        array is always returned, even if it contains only one byte.
    .Parameter String
        A string containing hex data in any of a variety of formats,
        including strings like the following, with or without extra
        tabs, spaces, quotes or other non-hex characters:
        0x41,0x42,0x43,0x44
        \x41\x42\x43\x44
        41-42-43-44
        41424344
        The string can be piped into the function too.
    .Link
        http://www.sans.org/windows-security/2010/02/11/powershell-byte-array-hex-convert
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [String] $String
    )
     
    #Clean out whitespaces and any other non-hex crud.
    #   Try to put into canonical colon-delimited format.
    #   Remove beginning and ending colons, and other detritus.
    $String = $String.ToLower() -replace '[^a-f0-9\\\,x\-\:]','' `
                                -replace '0x|\\x|\-|,',':' `
                                -replace '^:+|:+$|x|\\',''
     
    #Maybe there's nothing left over to convert...
    if ($String.Length -eq 0) { ,@() ; return } 
     
    #Split string with or without colon delimiters.
    if ($String.Length -eq 1) { 
        ,@([System.Convert]::ToByte($String,16))
    }
    elseif (($String.Length % 2 -eq 0) -and ($String.IndexOf(":") -eq -1)) { 
        ,@($String -split '([a-f0-9]{2})' | foreach-object {
            if ($_) {
                [System.Convert]::ToByte($_,16)
            }
        }) 
    }
    elseif ($String.IndexOf(":") -ne -1) { 
        ,@($String -split ':+' | foreach-object {[System.Convert]::ToByte($_,16)})
    }
    else { 
        ,@()
    }
    #The strange ",@(...)" syntax is needed to force the output into an
    #array even if there is only one element in the output (or none).
}



function ConvertTo-String($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++
    }  
    return $Val1Text  
}



function ConvertFrom-FolderID {
    <#
    .SYNOPSIS
    Convert Encoded Folder ID to Folder Path
     
    .PARAMETER EmailAddress
    The email address of the mailbox in question.  Can also be used as the return
    value from ConvertFrom-MailboxID
     
    .PARAMETER FolderID
    The mailbox identification string as provided by the DMS System
     
    .PARAMETER ImpersonationCredential
    The credential to use when accessing Exchange Web Services.
     
    .DESCRIPTION
    Takes the encoded Folder ID from the DMS System and returns the folder path for
    the Folder ID with the user mailbox.
     
    .EXAMPLE
    PS C:\> ConvertFrom-FolderID -EmailAddress "hubert.farnsworth@planetexpress.com" -FolderID "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000053414E4445584D42583031002F4F3D50697065722026204D6172627572792F4F553D504D2F636E3D526563697069656E74732F636E3D616265636B737465616400" -ImpersonationCredential $EWSAdmin
     
    \Inbox\Omicron Persei 8\Lrrr

    .EXAMPLE 
    PS C:\> $EmailAddress = ConvertFrom-MailboxID -MailboxID "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000042414C5445584D42583033002F4F3D50697065722026204D6172627572792F4F553D504D2F636E3D526563697069656E74732F636E3D6162313836353600D83521F3C10000000100000014000000850000002F6F3D50697065722026204D6172627572792F6F753D45786368616E67652041646D696E6973747261746976652047726F7570202846594449424F484632335350444C54292F636E3D436F6E66696775726174696F6E2F636E3D536572766572732F636E3D42414C5445584D4258303300420041004C005400450058004D0042005800300033002E00500069007000650072002E0052006F006F0074002E004C006F00630061006C0000000000"
    PS C:\> ConvertFrom-FolderID -EmailAddress $EmailAddress -FolderID "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000053414E4445584D42583031002F4F3D50697065722026204D6172627572792F4F553D504D2F636E3D526563697069656E74732F636E3D616265636B737465616400" -ImpersonationCredential $EWSAdmin
     
    \Inbox\Amphibios 9\Kif Kroker
     
    .NOTES
    This function requires Exchange Web Services Managed API version 1.2.
    The EWS Managed API can be obtained from: http://www.microsoft.com/en-us/download/details.aspx?id=28952
    #>
    [CmdletBinding()]
    param(
    	[Parameter(Mandatory=$true)]
    	[object]$EWSService,
     	[Parameter(Mandatory=$true)]
    	[string]$EmailAddress,
        [Parameter(Mandatory=$true)]
    	[string]$FolderID,
    	[Parameter(Mandatory=$false)]
    	[ValidateSet("EwsLegacyId", "EwsId", "EntryId", "HexEntryId", "StoreId", "OwaId")]
    	[string]$InputFormat = "EwsId",
    	[Parameter(Mandatory=$false)]
    	[ValidateSet("FolderPath", "EwsLegacyId", "EwsId", "EntryId", "HexEntryId", "StoreId", "OwaId")]
    	[string]$OutputFormat = "FolderPath"
    )
	Write-Verbose "Converting $FolderID from $InputFormat to $OutputFormat"
 
    #region Build Alternative ID Object
    $AlternativeIdItem  = New-Object Microsoft.Exchange.WebServices.Data.AlternateId
	$AlternativeIdItem.Mailbox = $EmailAddress
	$AlternativeIdItem.UniqueId = $FolderID
	$AlternativeIdItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::$InputFormat
    #endregion Build Alternative ID Object
 
    #region Retrieve Folder Path from EWS
    try {
        if ( $OutputFormat -eq "FolderPath" ) {
			# Build the Folder Property Set and then add Properties that we want
			$psFolderPropertySet = New-Object -TypeName Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
 
			# Define the Folder Extended Property Set Elements
			$PR_Folder_Path = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
 
			# Add to the Folder Property Set Collection
			$psFolderPropertySet.Add($PR_Folder_Path)
 
			$EwsFolderID = $EWSService.ConvertId($AlternativeIdItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)
	        $EwsFolder = New-Object Microsoft.Exchange.WebServices.Data.FolderID($EwsFolderID.UniqueId)
	        $TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($EWSService, $EwsFolder, $psFolderPropertySet)
	        
            # Retrieve the first Property (Folder Path in a Raw State)
	        $FolderPathRAW = $TargetFolder.ExtendedProperties[0].Value
	        # The Folder Path attribute actually contains non-ascii characters in place of the backslashes
	        #   Since the first character is one of these non-ascii characters, we use that for the replace method
	        $ConvertedFolderId = $FolderPathRAW.Replace($FolderPathRAW[0], "\")
		}
		else {
			$EwsFolderID = $Service.ConvertId($AlternativeIdItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::$OutputFormat )
			$ConvertedFolderId = $EwsFolderId.UniqueId
		}
    }
    catch {
        $ConvertedFolderId = $null
    }
    finally {
        $ConvertedFolderId
    }
    #endregion Retrieve Folder Path from EWS
}



function ConvertFrom-MailboxID {
    <#
    .SYNOPSIS
    Convert Encoded Mailbox ID to Email Address
     
    .PARAMETER MailboxID 
    The mailbox identification string as provided by the DMS System
     
    .DESCRIPTION
    Takes the encoded Mailbox ID from the DMS System and returns the email address of the end user.

    .EXAMPLE     
    PS C:\> ConvertFrom-MailboxID -MailboxID "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000000000001B55FA20AA6611CD9BC800AA002FC45A0C00000053414E4445584D42583031002F4F3D50697065722026204D6172627572792F4F553D504D2F636E3D526563697069656E74732F636E3D616265636B737465616400"
     
    John.Zoidberg@planetexpress.com

    .NOTES
    Requires active connection to the Active Directory infrastructure
    #>
    [CmdletBinding()]
    param(
     	[Parameter(Position=0,Mandatory=$true)]
    	[string]$MailboxID
    )
    try {
        $MailboxDN = ConvertTo-MailboxID -EncodedString $MailboxID 
        $ADSISearch = [DirectoryServices.DirectorySearcher]""
        $ADSISearch.Filter = "(&(&(&(objectCategory=user)(objectClass=user)(legacyExchangeDN=" + $MailboxDN + "))))"
        $SearchResults = $ADSISearch.FindOne()
        if ( -not $SearchResults ) {
            $ADSISearch.Filter = "(&(objectclass=user)(objectcategory=person)(proxyaddresses=x500:" + $MailboxDN + "))"
            $SearchResults = $ADSISearch.FindOne()    
        }
        $SearchResults.Properties.mail
    }
    catch {
        throw
    }
}



function ConvertTo-HexId {    
	param (
	        $EWSid,
            $EmailAddress
		  )
	process{
	    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
	    $aiItem.Mailbox = $EmailAddress
	    $aiItem.UniqueId = $EWSid   
	    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EWSId   
	    $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId) 
		return $convertedId.UniqueId
	}
}



function ConvertTo-MailboxID {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
        [String] $EncodedString
    )
 
	$ByteArray   = Convert-HexStringToByteArray -String $EncodedString
	$ByteArray   = $ByteArray | Where-Object { ( ($_ -ge 32) -and ($_ -le 127) ) -or ($_ -eq 0) }
	$ByteString  = Convert-ByteArrayToString -ByteArray $ByteArray -Encoding ASCII
	$StringArray = $ByteString.Split([char][int](0))
	$StringArray[21]
}



function Get-CallerPreference {
    <#
    .Synopsis
       Fetches "Preference" variable values from the caller's scope.
    .DESCRIPTION
       Script module functions do not automatically inherit their caller's variables, but they can be
       obtained through the $PSCmdlet variable in Advanced Functions.  This function is a helper function
       for any script module Advanced Function; by passing in the values of $ExecutionContext.SessionState
       and $PSCmdlet, Get-CallerPreference will set the caller's preference variables locally.
    .PARAMETER Cmdlet
       The $PSCmdlet object from a script module Advanced Function.
    .PARAMETER SessionState
       The $ExecutionContext.SessionState object from a script module Advanced Function.  This is how the
       Get-CallerPreference function sets variables in its callers' scope, even if that caller is in a different
       script module.
    .PARAMETER Name
       Optional array of parameter names to retrieve from the caller's scope.  Default is to retrieve all
       Preference variables as defined in the about_Preference_Variables help file (as of PowerShell 4.0)
       This parameter may also specify names of variables that are not in the about_Preference_Variables
       help file, and the function will retrieve and set those as well.
    .EXAMPLE
       Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

       Imports the default PowerShell preference variables from the caller into the local scope.
    .EXAMPLE
       Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name 'ErrorActionPreference','SomeOtherVariable'

       Imports only the ErrorActionPreference and SomeOtherVariable variables into the local scope.
    .EXAMPLE
       'ErrorActionPreference','SomeOtherVariable' | Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

       Same as Example 2, but sends variable names to the Name parameter via pipeline input.
    .INPUTS
       String
    .OUTPUTS
       None.  This function does not produce pipeline output.
    .LINK
       about_Preference_Variables
    #>

    [CmdletBinding(DefaultParameterSetName = 'AllVariables')]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ $_.GetType().FullName -eq 'System.Management.Automation.PSScriptCmdlet' })]
        $Cmdlet,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.SessionState]$SessionState,

        [Parameter(ParameterSetName = 'Filtered', ValueFromPipeline = $true)]
        [string[]]$Name
    )

    begin {
        $filterHash = @{}
    }
    
    process {
        if ($null -ne $Name)
        {
            foreach ($string in $Name)
            {
                $filterHash[$string] = $true
            }
        }
    }

    end {
        # List of preference variables taken from the about_Preference_Variables help file in PowerShell version 4.0

        $vars = @{
            'ErrorView' = $null
            'FormatEnumerationLimit' = $null
            'LogCommandHealthEvent' = $null
            'LogCommandLifecycleEvent' = $null
            'LogEngineHealthEvent' = $null
            'LogEngineLifecycleEvent' = $null
            'LogProviderHealthEvent' = $null
            'LogProviderLifecycleEvent' = $null
            'MaximumAliasCount' = $null
            'MaximumDriveCount' = $null
            'MaximumErrorCount' = $null
            'MaximumFunctionCount' = $null
            'MaximumHistoryCount' = $null
            'MaximumVariableCount' = $null
            'OFS' = $null
            'OutputEncoding' = $null
            'ProgressPreference' = $null
            'PSDefaultParameterValues' = $null
            'PSEmailServer' = $null
            'PSModuleAutoLoadingPreference' = $null
            'PSSessionApplicationName' = $null
            'PSSessionConfigurationName' = $null
            'PSSessionOption' = $null

            'ErrorActionPreference' = 'ErrorAction'
            'DebugPreference' = 'Debug'
            'ConfirmPreference' = 'Confirm'
            'WhatIfPreference' = 'WhatIf'
            'VerbosePreference' = 'Verbose'
            'WarningPreference' = 'WarningAction'
        }

        foreach ($entry in $vars.GetEnumerator()) {
            if (([string]::IsNullOrEmpty($entry.Value) -or -not $Cmdlet.MyInvocation.BoundParameters.ContainsKey($entry.Value)) -and
                ($PSCmdlet.ParameterSetName -eq 'AllVariables' -or $filterHash.ContainsKey($entry.Name))) {
                
                $variable = $Cmdlet.SessionState.PSVariable.Get($entry.Key)
                
                if ($null -ne $variable) {
                    if ($SessionState -eq $ExecutionContext.SessionState) {
                        Set-Variable -Scope 1 -Name $variable.Name -Value $variable.Value -Force -Confirm:$false -WhatIf:$false
                    }
                    else {
                        $SessionState.PSVariable.Set($variable.Name, $variable.Value)
                    }
                }
            }
        }

        if ($PSCmdlet.ParameterSetName -eq 'Filtered') {
            foreach ($varName in $filterHash.Keys) {
                if (-not $vars.ContainsKey($varName)) {
                    $variable = $Cmdlet.SessionState.PSVariable.Get($varName)
                
                    if ($null -ne $variable)
                    {
                        if ($SessionState -eq $ExecutionContext.SessionState)
                        {
                            Set-Variable -Scope 1 -Name $variable.Name -Value $variable.Value -Force -Confirm:$false -WhatIf:$false
                        }
                        else
                        {
                            $SessionState.PSVariable.Set($variable.Name, $variable.Value)
                        }
                    }
                }
            }
        }
    }
}



function Get-EWSUserDN {
    [CmdletBinding()] 
    param(
        [parameter(HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [Parameter(Position=1, Mandatory=$true)]
        [string]$EmailAddress
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
    $ExchangeVersion = [ews_exchver]::$($EWSService.RequestedServerVersion)
    $adService = New-Object ews_autod($ExchangeVersion)
    $adService.Credentials = $EWSService.Credentials
    $adService.EnableScpLookup = $false;
    $adService.RedirectionUrlValidationCallback = {$true}
    $UserSettings = new-object ews_usersettingname[] 1
    $UserSettings[0] = [ews_usersettingname]::UserDN
    $adResponse = $adService.GetUserSettings($EmailAddress, $UserSettings);
    return $adResponse.Settings[[ews_usersettingname]::UserDN]
}



function Get-ScriptPath {
	$scriptDir = Get-Variable PSScriptRoot -ErrorAction SilentlyContinue | ForEach-Object { $_.Value }
	if (!$scriptDir) {
		if ($MyInvocation.MyCommand.Path) {
			$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
		}
	}
	if (!$scriptDir) {
		if ($ExecutionContext.SessionState.Module.Path) {
			$scriptDir = Split-Path (Split-Path $ExecutionContext.SessionState.Module.Path)
		}
	}
	if (!$scriptDir) {
		$scriptDir = $PWD
	}
	
	return $scriptDir
}



function Get-WebFile {
    <#
    .SYNOPSIS
    Downloads file from the web and returns the full path when complete.

    .DESCRIPTION
    Downloads file from the web and returns the full path when complete.

    .EXAMPLE
    $source = 'http://download.microsoft.com/download/3/E/4/3E4AF215-E418-47B8-BB89-D5555E858728/EwsManagedApi.MSI'
    Get-WebFile -source $source
    
    Description
    -----------
    Downloads the EWS managed api installer to a temporary directory and returns the final location when completed.
    #>
    [CmdLetBinding()]
    param(
        [string]$source,
        [string]$destination
    )
    if ([string]::IsNullOrEmpty($destination)) {
        $TempDirPath = "$($Env:TEMP)\$([System.Guid]::NewGuid().ToString())"
        Write-Verbose "$($MyInvocation.MyCommand): Creating temporary directory $TempDirPath"
        [string]$NewDir = New-Item -Type Directory -Path $TempDirPath
        $filename = $source.Split('/') | Select -Last 1
        $destfullpath = $NewDir + '\' + $filename
    }
    elseif (Test-Path (Split-Path $destination -Parent)){
        $destfullpath = $destination
    }
    else {
        throw '$($MyInvocation.MyCommand): Unable to validate destination path exists!'
    }
    try {
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($source, $destfullpath)
        return $destfullpath
    }
    catch {
        throw '$($MyInvocation.MyCommand): Unable to download file!'
    }
}



function Invoke-MSIExec {
    <#
    .SYNOPSIS
    Invokes msiexec.exe

    .DESCRIPTION
    Runs msiexec.exe, passing all the arguments that get passed to `Invoke-MSIExec`.

    .EXAMPLE
    Invoke-MSIExec /a C:\temp\EwsManagedApi.MSI /qb TARGETDIR=c:\Scripts\EWSAPI

    Runs `/a C:\temp\EwsManagedApi.MSI /qb TARGETDIR=c:\temp\EWSAPI`, which extracts the contents of C:\temp\EwsManagedApi.MSI into c:\temp\EWSAPI
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromRemainingArguments=$true)]
        $Args
    )
    
    Write-Verbose ($Args -join " ")
    # Note: The Out-Null forces the function to wait until the prior process completes, nifty
    & (Join-Path $env:SystemRoot 'System32\msiexec.exe') $Args | Out-Null
}



function New-UniqueFileName {
    param(
        [Parameter(Position=0, Mandatory=$true)] [string]$FileName
    )
    Begin
    {
    
    $directoryName = [System.IO.Path]::GetDirectoryName($FileName)
    $FileDisplayName = [System.IO.Path]::GetFileNameWithoutExtension($FileName);
    $FileExtension = [System.IO.Path]::GetExtension($FileName);
    for ($i = 1; ; $i++){
            
            if (![System.IO.File]::Exists($FileName)){
                return($FileName)
            }
            else{
                    $FileName = [System.IO.Path]::Combine($directoryName, $FileDisplayName + "(" + $i + ")" + $FileExtension);
            }                
            
            if($i -eq 10000){throw "Out of Range"}
        }
    }
}



function Test-EmailAddressFormat {
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='String to validate email address format.')]
        [string]$emailaddress
    )
    $emailregex = "[a-z0-9!#\$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#\$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
    if ($emailaddress -imatch $emailregex ) {
        return $true
    }
    else {
        return $false
    }
}



function Test-UserSIDFormat {
    [CmdletBinding()]
    param(
        [parameter(Position=0, Mandatory=$True, HelpMessage='String to validate is in user SID format.')]
        [string]$SID
    )
    $sidregex = "^S-\d-\d+-(\d+-){1,14}\d+$"
    if ($SID -imatch $sidregex ) {
        return $true
    }
    else {
        return $false
    }
}



Function Validate-EmailAddres {
    param( 
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress
    )
    try {
        $check = New-Object System.Net.Mail.MailAddress($EmailAddress)
        return $true
    }
    catch {
        return $false
    }
}



## PUBLIC MODULE FUNCTIONS AND DATA ##

function Connect-EWS {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Connect-EWS.md
    #>
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
    
    # If ServerCertificateValidationCallback is set to anything you will experience all kinds of issues so we temporarily nullify it for the session.
    # I'm open to alternatives that work on this one....
    $tempCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
    if ($tempCallback -ne $null) {
        Write-Warning "$($FunctionName): ServerCertificateValidationCallback being set to null for this session."
        Set-ServerCertificateValidationCallback [System.Net.ServicePointManager]::ServerCertificateValidationCallback
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
    }

    return $true
}




function Get-EmailAddressFromAD {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EmailAddressFromAD.md
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='ID to lookup. Defaults to current users SID')]
        [string]$UserID = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    )

    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
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
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSCalendarAppointments.md
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

    Write-Verbose "$($FunctionName): Attempting to gather calendar entries for $($email)"

    $MailboxToAccess = New-Object ews_mailbox($email)

    if ([string]::IsNullOrEmpty($FolderPath)) {
        $FolderID = New-Object ews_folderid([ews_wellknownfolder]::Calendar, $MailboxToAccess)
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
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSCalenderViewAppointments.md
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




function Get-EWSContact {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSContact.md
    #>
    [CmdletBinding()] 
    param(
        [parameter(Position = 0)]
        [ews_service]$EWSService,
        [parameter(Position = 1)]
        [string]$Mailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [string]$EmailAddress,
        [Parameter(Position = 3)]
        [string]$Folder,
        [Parameter(Position = 4)]
        [ValidateSet('DirectoryOnly','DirectoryThenContacts','ContactsOnly','ContactsThenDirectory')]
        [string]$SearchType = 'ContactsThenDirectory',
        [Parameter(Position=5)]
        [switch]$Partial
    )  
    # Pull in all the caller verbose,debug,info,warn and other preferences
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = MyInvocation.MyCommand.Name
    
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

    if( -not [string]::IsNullOrEmpty($Folder) ) {
        $folderid = Get-EWSFolder -EWSService $EWSService -FolderPath $Folder -Mailbox $Mailbox
    }
    else {
        $folderid = Get-EWSFolder -EWSService $EWSService -FolderObject Contacts -Mailbox $Mailbox
    }
    
    # Get the parent folder ID
    $type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
    $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
    $ParentFolderIds = [Activator]::CreateInstance($type)
    $ParentFolderIds.Add($folderid.Id)

    try {
        $ncCol = @($EWSService.ResolveName($EmailAddress,$ParentFolderIds,$SearchType,$true))
    }
    catch {
        throw "$($FunctionName): Unable to resolve contact $EmailAddress"
    }
    foreach($Result in $ncCol) {
        # If the Contact property is null then this is likely just a user contact
        if($Result.Contact -eq $null) {
            if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -or $Partial){
                # Convert to a full EWS contact and return
                Write-Verbose "$($FunctionName): Returning contact from the mailbox contacts."
                [ews_contact]::Bind($EWSService,$Result.Mailbox.Id)
            }
        }
        # Otherwise it was probably found in the directory.
        else{
            if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -or $Partial){
                if($Result.Mailbox.MailboxType -eq [ews_mailboxtype]::Mailbox){
                    $UserDn = Get-EWSUserDN -EWSService $EWSService -EmailAddress $Result.Mailbox.Address
                    $ncCola = $EWSService.ResolveName($UserDn,$ParentFolderIds,$SearchType,$true)
                    Write-Verbose ("$($FunctionName): Number of matching Contacts Found - " + $ncCola.Count)
                    foreach($aResult in $ncCola){
                        Write-Verbose "$($FunctionName): Returning contact from the directory."
                        $aResult.Contact
                    }
                }
            }
        }
    }
}




function Get-EWSContacts {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSContacts.md
    #>
    [CmdletBinding()] 
    param(
        [parameter(HelpMessage = 'Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position = 1, HelpMessage = 'Mailbox to target.')]
        [string]$Mailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [string]$EmailAddress,
        [Parameter(Position = 3)]
        [string]$Folder,
        [Parameter(Position = 4)]
        [ValidateSet('DirectoryOnly','DirectoryThenContacts','ContactsOnly','ContactsThenDirectory')]
        [ews_resolvenamelocation]$SearchType = 'ContactsThenDirectory',
        [Parameter(Position = 5)]
        [switch]$Partial
    )

    # Pull in all the caller verbose,debug,info,warn and other preferences
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = MyInvocation.MyCommand.Name
    
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

    if( -not [string]::IsNullOrEmpty($Folder) ) {
        $folderid = Get-EWSFolder -EWSService $EWSService -FolderPath $Folder -Mailbox $Mailbox
    }
    else {
        $folderid = Get-EWSFolder -EWSService $EWSService -FolderObject Contacts -Mailbox $Mailbox
    }

    $SfSearchFilter = New-Object ews_searchfilter_isequalto([ews_schema_item]::ItemClass,"IPM.Contact")
    $ivItemView =  New-Object ews_itemview(1000)
    $fiItems = $null

    do {
        $fiItems = $EWSService.FindItems($folderid.Id,$SfSearchFilter,$ivItemView)
        if($fiItems.Items.Count -gt 0) {
            $psPropset = new-object ews_propset([ews_basepropset]::FirstClassProperties)
            [Void]$EWSService.LoadPropertiesForItems($fiItems,$psPropset)
            foreach($Item in $fiItems.Items) {
                Write-Output $Item
            }
        }
        $ivItemView.Offset += $fiItems.Items.Count    
    } while($fiItems.MoreAvailable -eq $true) 
}




function Get-EWSDllLoadState {
   <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSDllLoadState.md
    #>
    if (-not (get-module Microsoft.Exchange.WebServices)) {
        return $false
    }
    else {
        return $true
    }
}




function Get-EWSFolder {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSFolder.md
    #>
    [CmdletBinding(DefaultParametersetName='FolderAsString')]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [parameter(ParameterSetName='FolderAsString')]
        [parameter(ParameterSetName='FolderAsObject')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox of folder.')]
        [parameter(ParameterSetName='FolderAsString')]
        [parameter(ParameterSetName='FolderAsObject')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Folder path.')]
        [parameter(ParameterSetName='FolderAsString')]
        [string]$FolderPath,
        [parameter(Position=2, HelpMessage='Well known folder object.')]
        [parameter(ParameterSetName='FolderAsObject')]
        [ValidateSet('Calendar','Contacts','DeletedItems','Drafts','Inbox','Journal','Notes','Outbox','SentItems','Tasks','MsgFolderRoot','PublicFoldersRoot','Root','JunkEmail','SearchFolders','VoiceMail','RecoverableItemsRoot','RecoverableItemsDeletions','RecoverableItemsVersions','RecoverableItemsPurges','ArchiveRoot','ArchiveMsgFolderRoot','ArchiveDeletedItems','ArchiveRecoverableItemsRoot','ArchiveRecoverableItemsDeletions','ArchiveRecoverableItemsVersions','ArchiveRecoverableItemsPurges','SyncIssues','Conflicts','LocalFailures','ServerFailures','RecipientCache','QuickContacts','ConversationHistory','ToDoSearch')]
        [ews_wellknownfolder]$FolderObject,
        [parameter(Position=3, HelpMessage='Are you targeting a public Folder Path?')]
        [parameter(ParameterSetName='FolderAsString')]
        [switch]$PublicFolder
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
    
    # Return a reference to a folder specified by path 
    
    switch ($PsCmdlet.ParameterSetName) { 
        'FolderAsString' {
            if ($PublicFolder) { 
                $mbx = ''
                try {
                    $Folder = [ews_folder]::Bind($EWSService, [ews_wellknownfolder]::PublicFoldersRoot) 
                }
                catch {
                    Write-Warning "$($FunctionName): Unable to find a public folder server or database to connect to."
                    return $null
                }
            }
            else {
                $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
                $mbx = New-Object ews_mailbox($email)
                $FolderID = New-Object ews_folderid([ews_wellknownfolder]::MsgFolderRoot, $mbx )
            }
            
            if ($FolderPath -ne '\') {
                $PathElements = $FolderPath -split '\\' 
                For ($i=0; $i -lt $PathElements.Count; $i++) { 
                    if ($PathElements[$i]) { 
                        $View = New-Object ews_folderview(2,0) 
                        $View.PropertySet = [ews_basepropset]::IdOnly
                        $SearchFilter = New-Object ews_searchfilter_isequalto([ews_schema_folder]::DisplayName, $PathElements[$i])
                        $FolderResults = $Folder.FindFolders($SearchFilter, $View) 
                        if ($FolderResults.TotalCount -ne 1) { 
                            # We have either none or more than one folder returned... Either way, we can't continue 
                            Write-Verbose "$($FunctionName): Failed to find $($PathElements[$i]), path requested was $FolderPath"
                            return $null
                        }
                         
                        if (-not [String]::IsNullOrEmpty(($mbx))) {
                            $folderId = New-Object ews_folderid($FolderResults.Folders[0].Id, $mbx) 
                            try {
                                $Folder = [ews_folder]::Bind($service, $folderId) 
                            }
                            catch {
                                Write-Warning "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
                                return $null
                            }
                        } 
                        else {
                            try {
                                $Folder = [ews_folder]::Bind($service, $FolderResults.Folders[0].Id)
                            }
                            catch {
                                Write-Warning "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
                                return $null
                            }
                        } 
                    } 
                } 
            }
            else {
                try {
                    $Folder = [ews_folder]::Bind($EWSService, $FolderID)
                }
                catch {
                    Write-Warning "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
                    return $null
                }
            }
        }
        'FolderAsObject' {
            $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
            $mbx = New-Object ews_mailbox($email)
            $FolderID = New-Object ews_folderid($FolderObject, $mbx)
            try {
                $Folder = [ews_folder]::Bind($EWSService, $FolderID)
            }
            catch {
                Write-Warning "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
                return $null
            }
        }
    }
    return $Folder 
}




function Get-EWSFolderPaths {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSFolderPaths.md
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [Parameter(Position=1, Mandatory=$true)] 
        [ews_folderid]$RootFolderId,
        [Parameter(Position=2, Mandatory=$true)]
        [PSObject]$FolderCache,
        [Parameter(Position=3)]
        [String]$FolderPrefix
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
    
    #Define Extended properties  
    $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)  
    $PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long)
    
    #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
    $fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
    
    #Deep transversal will ensure all folders in the search path are returned  
    $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
    
    #Add Properties to the  Property Set  
    $psPropertySet.Add($PR_Folder_Path)
    $psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED)
    $fvFolderView.PropertySet = $psPropertySet 
    
    #The Search filter will exclude any Search Folders  
    $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")
    $fiResult = $null
    
    #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
    do {  
        $fiResult = $EWSService.FindFolders($RootFolderId,$sfSearchFilter,$fvFolderView)  
        foreach($ffFolder in $fiResult.Folders) {
            #Try to get the FolderPath Value and then covert it to a usable String 
            $foldpathval = $null
            if ($ffFolder.TryGetProperty($PR_Folder_Path,[ref]$foldpathval)) {  
                $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
                $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
                $hexString = $hexArr -join ''  
                $hexString = $hexString.Replace("FEFF", "5C00")  
                $fpath = ConvertTo-String($hexString)  
            }
            if($FolderCache.ContainsKey($ffFolder.Id.UniqueId) -eq $false) {
                if ([string]::IsNullOrEmpty($FolderPrefix)) {
                    $FolderCache.Add($ffFolder.Id.UniqueId,($fpath))    
                }
                else {
                    $FolderCache.Add($ffFolder.Id.UniqueId,("\" + $FolderPrefix + $fpath))    
                }
            }
        } 
        $fvFolderView.Offset += $fiResult.Folders.Count
    } while($fiResult.MoreAvailable)
}




function Get-EWSModuleInitializationState {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSModuleInitializationState.md
    #>
    return $script:EWSModuleInitialized
}




function Get-EWSOOFSettings {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSOofSettings.md
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position=1, Mandatory=$True, ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [String]$Mailbox
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
            Mailbox = $TargetedMailbox
        }
    }
    catch {
        throw "$($FunctionName): Unable to get out of office info for $TargetedMailbox"
    }
}




function Get-EWSService {
     <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSService.md
    #>
    return $script:modEWSService
}




function Get-EWSTargettedMailbox {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-EWSTargetedMailbox.md
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox you are targeting.')]
        [string]$Mailbox
    )
    
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
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
                throw "$($FunctionName): Unable to get a mailbox for this account from AD. Ensure you are running this from a domain joined computer."
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




function Get-ServerCertificateValidationCallback {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Get-ServerCertificateValidationCallback.md
    #>
    return $script:modCertCallback
}




function Import-EWSDll {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Import-EWSDll.md
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=0)]
        [string]$EWSManagedApiPath
    )
    
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
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




function Initialize-EWS {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Initialize-EWS.md
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=0)]
        [string]$EWSManagedApiPath,
        [parameter(Position=1)]
        [switch]$Uninitialize
    )

    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
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
            $accelerators = [PSObject].Assembly.GetType('System.Management.Automation.TypeAccelerators')::get
            $accelkeyscopy = @{}
            $accelerators.Keys | Where {$_ -like "ews_*"} | Foreach { $accelkeyscopy.$_ = $EWSAccels[$_] }
            foreach ( $key in $accelkeyscopy.Keys ) {
                Write-Verbose "UnInitialize-EWS: Removing type accelerator - $($key)"
                $accelerators.Remove($key) | Out-Null
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




function Install-EWSDll {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Install-EWSDll.md
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




function New-EWSCalendarEntry {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/New-EWSCalendarEntry.md
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




function New-EWSExtendedProperty {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/New-EWSExtendedProperty.md
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




function Remove-EWSCalendarAppointment {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Remove-EWSCalendarAppointment.md
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(Position=1, Mandatory=$true,ValueFromPipeline=$true, HelpMessage='Calendar appointment object to remove.')]
        [Microsoft.Exchange.WebServices.Data.Appointment]$Appointment,
        [parameter(Position=2, HelpMessage='Deletion mode.')]
        [ValidateSet('HardDelete','SoftDelete','MoveToDeletedItems')]
        [string]$DeleteMode = 'HardDelete',
        [parameter(Position=3, HelpMessage='Cancellation mode.')]
        [ValidateSet('SendToNone','SendOnlyToAll','SendOnlyToChanged','SendToAllAndSaveCopy','SendToChangedAndSaveCopy')]
        [string]$CancellationMode = 'SendToNone'
    )
   Begin {
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
   }
    Process {
         $Appointment | Foreach {
            Write-Verbose "$($FunctionName): Deleting existing appointment with the subject of $($_.Subject)"
            try {                
                $_.Delete([ews_deletemode]::$DeleteMode,[ews_sendcancellationmode]::$CancellationMode)
            }
            catch {
                Write-Warning "$($FunctionName): Unable to DELETE existing appointment!"
            }
         }
    }
    End {}
}




function Set-EWSMailboxImpersonation {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Set-EWSMailboxImpersonation.md
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(Position=1, Mandatory=$True, HelpMessage='Mailbox to impersonate.')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Do not attempt to validate rights against this mailbox (can speed up operations)')]
        [switch]$SkipValidation
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




function Set-EWSModuleInitializationState {
     <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Set-EWSModuleInitializationState.md
    #>
    param(
        [bool]$State = $false
    )
    $script:EWSModuleInitialized = $State
}




function Set-EWSOofSettings {
     <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Set-EWSOofSettings.md
    #>
    [CmdletBinding(DefaultParameterSetName='Default')]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.', ParameterSetName = 'Default')]
        [parameter(Position=0, HelpMessage='Connected EWS object.', ParameterSetName = 'Enabled')]
        [parameter(Position=0, HelpMessage='Connected EWS object.', ParameterSetName = 'Disabled')]
        [parameter(Position=0, HelpMessage='Connected EWS object.', ParameterSetName = 'Scheduled')]
        [ews_service]$EWSService,
        [parameter(Position=1, Mandatory=$True, ParameterSetName = 'Default', ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [parameter(Position=1, Mandatory=$True, ParameterSetName = 'Enabled', ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [parameter(Position=1, Mandatory=$True, ParameterSetName = 'Disabled', ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [parameter(Position=1, Mandatory=$True, ParameterSetName = 'Scheduled', ValueFromPipelineByPropertyName=$true, HelpMessage='Mailbox to impersonate.')]
        [String]$Mailbox,
        [Parameter(Position=2, ParameterSetName = 'Default')]
        [Parameter(Position=2, ParameterSetName = 'Enabled')]
        [Parameter(Position=2, ParameterSetName = 'Disabled')]
        [Parameter(Position=2, ParameterSetName = 'Scheduled')]
        [ValidateSet("Enabled","Disabled","Scheduled")]
        [String]$State = 'Disabled',
        [Parameter(Position=3, ParameterSetName = 'Enabled')]
        [Parameter(Position=3, ParameterSetName = 'Scheduled')]
        [ValidateSet("All","External","None")]
        [String]$ExternalAudience = 'All',
        [Parameter(Position=4, ParameterSetName = 'Enabled')]
        [Parameter(Position=4, ParameterSetName = 'Scheduled')]
        [DateTime]$StartTime,
        [Parameter(Position=5, ParameterSetName = 'Enabled')]
        [Parameter(Position=5, ParameterSetName = 'Scheduled')]
        [DateTime]$EndTime,        
        [Parameter(Position=6, ParameterSetName = 'Enabled')]
        [Parameter(Position=6, ParameterSetName = 'Scheduled')]
        [String]$InternalReply,
        [Parameter(Position=7, ParameterSetName = 'Enabled')]
        [Parameter(Position=7, ParameterSetName = 'Scheduled')]
        [String]$ExternalReply
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
        if (($p.key -ne "Mailbox") -and ($p.key -ne "EWSService"))  {
            $oof."$($p.key)" = $p.value
        }
    }

    $oof.State = [Microsoft.Exchange.WebServices.Data.OofState]::$State
    $EWSService.SetUserOofSettings($Mailbox,$oof)
}




function Set-EWSService {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Set-EWSService.md
    #>
    param(
        [ews_service]$ConnectedService = $null
    )

    $script:modEWSService = $ConnectedService
}




Function Set-EWSSSLIgnoreWorkaround {
      <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Set-EWSSSLIgnoreWorkaround.md
    #>
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




function Set-ServerCertificateValidationCallback {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Set-ServerCertificateValidationCallback.md
    #>

    param(
        [string]$CertCallback = [System.Net.ServicePointManager]::ServerCertificateValidationCallback
    )

    $script:modCertCallback = $CertCallback
}




function Test-EWSAutodiscover {
    <#
    .EXTERNALHELP EWSModule-help.xml
    .LINK
        https://github.com/zloeber/EWSModule/tree/master/release/0.0.1/docs/Test-EWSAutodiscover.md
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




## Post-Load Module code ##


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


#region Module Cleanup
$ExecutionContext.SessionState.Module.OnRemove = {
        if ( Initialize-EWS -Uninitialize ) {}
        else { Write-Warning "Unable to uninitialize module" }
}

$null = Register-EngineEvent -SourceIdentifier ( [System.Management.Automation.PsEngineEvent]::Exiting ) -Action {
        if ( Initialize-EWS -Uninitialize ) {}
        else { Write-Warning "Unable to uninitialize module" }
}
#endregion Module Cleanup




