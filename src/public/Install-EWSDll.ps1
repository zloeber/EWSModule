function Install-EWSDll {
    <#
    .SYNOPSIS
    Attempts to download and extract the ews dll needed for this library.

    .DESCRIPTION
    Attempts to download and extract the ews dll needed for this library.
    #>
    [CmdLetBinding()]
    param(
        [string]$source = 'http://download.microsoft.com/download/8/9/9/899EEF2C-55ED-4C66-9613-EE808FCF861C/EwsManagedApi.msi',
        [string]$destination,
        [switch]$SkipDownload
    )
    
    if ([string]::IsNullOrEmpty($destination) -or (-not (Test-Path $destination))) {
        $destination = (Split-Path (get-variable myinvocation -scope script).value.Mycommand.Definition -Parent) + "\EwsManagedApi.MSI"
    }
    if (-not $SkipDownload) {
        try {
            $splatparam = @{
                'source' = $source
                'destination' = $destination
            }
            Write-Verbose 'Install-EWSDll: Attempting to download MSI.'
            $Download = Get-WebFile @splatparam
        }
        catch {
            throw 'Install-EWSDll: Unable to download the EWS install file!'
        }
        $DestPath = (Split-Path $Download -Parent) + "\EWSFiles\"
    }
    else {
        $DestPath = ".\EWSFiles\"
    }
    Write-Verbose "Install-EWSDll: Attempting to extract MSI to $($DestPath)."
    Invoke-MSIExec /quiet /a $Download /qn TARGETDIR=$DestPath

    if (Test-Path ("$DestPath\Microsoft.Exchange.WebServices.dll")) {
        Write-Verbose "Install-EWSDll: Copying MSI back to the download path."
        Copy-Item -Path "$DestPath\Microsoft.Exchange.WebServices.dll" -Destination (Split-Path $Download -Parent)
        Remove-Item -Recurse -Force -Path $DestPath
        Remove-Item -Force $Download
    }
    else {
        throw 'Install-EWSDll: Unable to extract the EWS install file!'
    }
}