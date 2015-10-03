﻿# Run this file to install the %%ModuleName%% PowerShell module with the following 
# in an administrative PowerShell prompt:
#
# 	iex (New-Object Net.WebClient).DownloadString("https://github.com/zloeber/EWSModule/raw/master/Install.ps1")

# Some general variables
$ModuleName = 'EWSModule'	# Example: mymodule
$GithubURL = 'https://github.com/zloeber/EWSModule'	# Example: https://www.github.com/zloeber/mymodule

# Download and install the module
$webclient = New-Object System.Net.WebClient
$url = "$GithubURL/archive/master.zip"
Write-Host "Downloading latest version of EWSModule from $url" -ForegroundColor Cyan
$file = "$($env:TEMP)\$($ModuleName).zip"
$webclient.DownloadFile($url,$file)
Write-Host "File saved to $file" -ForegroundColor Green
$targetondisk = "$($env:USERPROFILE)\Documents\WindowsPowerShell\Modules"
New-Item -ItemType Directory -Force -Path $targetondisk | out-null
$shell_app=new-object -com shell.application
$zip_file = $shell_app.namespace($file)
Write-Host "Uncompressing the Zip file to $($targetondisk)" -ForegroundColor Cyan
$destination = $shell_app.namespace($targetondisk)
$destination.Copyhere($zip_file.items(), 0x10)
Write-Host "Renaming folder" -ForegroundColor Cyan
if (Test-Path "$targetondisk\$($ModuleName)") { Remove-Item -Force "$targetondisk\$($ModuleName)" -Confirm:$false }
Rename-Item -Path ($targetondisk+"\$($ModuleName)-master") -NewName "$ModuleName" -Force
Write-Host "Module has been installed" -ForegroundColor Green
Write-Host "You can now import the module with: Import-Module -Name $ModuleName"