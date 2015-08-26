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