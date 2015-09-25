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