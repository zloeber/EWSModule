function Connect-EWS {
    <#
    .SYNOPSIS
        Connects to Exchange Web Services.
    .DESCRIPTION
        Connects to Exchange Web Services. Allows for multiple methods to connect, including autodiscover.
        Note that your login ID must be in email format for autodiscover to function.
    .PARAMETER UserName
        Username to connect with.
    .PARAMETER Password
        Password to connect with.
    .PARAMETER Domain
        Domain to connect to.
    .PARAMETER Credential
        Credential object to connect with.
    .PARAMETER ExchangeVersion
        Version of Exchange to target.
    .PARAMETER EWSUrl
        Exchange web services url to connect to.
    .PARAMETER EWSTracing
        Enable EWS tracing.
    .PARAMETER IgnoreSSLCertificate
        Ignore SSL validation checks.

    .EXAMPLE
       PS > $credentials = Get-Credential
       PS > Connect-EWS -Creds $credentials -ExchangeVersion 'Exchange2013_SP1' -EwsUrl 'https://webmail.contoso.com/ews/Exchange.asmx'
       
       Description
       -----------
       Connects to Exchange web services with credentials provided at the prompt.

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
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

    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand

    # If the things were not initialized manually then this will help eliminate one step
    $null = Initialize-EWS

    if (Get-EWSModuleInitializationState) {

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
    else {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
}