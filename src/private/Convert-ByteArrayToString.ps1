﻿function Convert-ByteArrayToString {
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
 