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