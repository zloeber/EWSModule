function Test-EWSAutodiscover {
    <#
    .SYNOPSIS
    This function uses the EWS Managed API to test the Exchange Autodiscover service.

    .DESCRIPTION
    This function will retreive the Client Access Server URLs for a specified email address
    by querying the autodiscover service of the Exchange server.

    .PARAMETER  EmailAddress
    Specifies the email address for the mailbox that should be tested.

    .PARAMETER  Location
    Set to External by default, but can also be set to Internal. This parameter controls whether
    the internal or external URLs are returned.

    .PARAMETER  Credential
    Specifies a user account that has permission to perform this action. Type a user name, such as 
    "User01" or "Domain01\User01", or enter a PSCredential object, such as one from the Get-Credential cmdlet.

    .PARAMETER  TraceEnabled
    Use this switch parameter to enable tracing. This is used for debugging the XML response from the server.    

    .PARAMETER  IgnoreSsl
    Set to $true by default. If you do not want to ignore SSL warnings or errors, set this parameter to $false.

    .PARAMETER  Url
    You can use this parameter to manually specifiy the autodiscover url.        

    .EXAMPLE
    PS C:\> Test-Autodiscover -EmailAddress administrator@uclabs.ms -Location internal

    This example shows how to retrieve the internal autodiscover settings for a user.

    .EXAMPLE
    PS C:\> Test-Autodiscover -EmailAddress administrator@uclabs.ms -Credential $cred

    This example shows how to retrieve the external autodiscover settings for a user. You can
    provide credentials if you do not want to use the Windows credentials of the user calling
    the function.

    .LINK
    http://msdn.microsoft.com/en-us/library/dd633699%28v=EXCHG.80%29.aspx

    .LINK
    http://www.the-little-things.net/

    .LINK
    https://www.github.com/zloeber/EWSModule

    .NOTES
    Author: Zachary Loeber
    Requires: Powershell 3.0
    Version History
    1.0.0 - Initial release
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

    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

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