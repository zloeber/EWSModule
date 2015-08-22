# The Convert-HexStringToByteArray and Convert-ByteArrayToString functions are from
# Link: http://www.sans.org/windows-security/2010/02/11/powershell-byte-array-hex-convert
function Convert-HexStringToByteArray {
    ################################################################
    #.Synopsis
    # Convert a string of hex data into a System.Byte[] array. An
    # array is always returned, even if it contains only one byte.
    #.Parameter String
    # A string containing hex data in any of a variety of formats,
    # including strings like the following, with or without extra
    # tabs, spaces, quotes or other non-hex characters:
    # 0x41,0x42,0x43,0x44
    # \x41\x42\x43\x44
    # 41-42-43-44
    # 41424344
    # The string can be piped into the function too.
    ################################################################
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
 
function Convert-ByteArrayToString {
    <#
    .Synopsis
    Returns the string representation of a System.Byte[] array. ASCII string is the default, but Unicode, UTF7, UTF8 and UTF32 are available too.
    .Parameter ByteArray
    System.Byte[] array of bytes to put into the file. If you pipe this array in, you must pipe the [Ref] to the array. Also accepts a single Byte object instead of Byte[].
    .Parameter Encoding
    Encoding of the string: ASCII, Unicode, UTF7, UTF8 or UTF32. ASCII is the default.
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
        $MailboxDN = ConvertTo-MailboxIdentification -EncodedString $MailboxID 
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
 
function ConvertTo-HexId{    
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