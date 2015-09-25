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