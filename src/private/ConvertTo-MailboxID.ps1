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