﻿function Convert-HexStringToByteArray {
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