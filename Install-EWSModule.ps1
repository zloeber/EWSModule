# Install EWSModule with the following:
# iex (New-Object Net.WebClient).DownloadString("https://github.com/zloeber/EWSModule/raw/master/Install-EWSModule.ps1")
$webclient = New-Object System.Net.WebClient
$url = "https://github.com/zloeber/EWSModule/archive/master.zip"
Write-Host "Downloading latest version of EWSModule from $url" -ForegroundColor Cyan
$file = "$($env:TEMP)\EWSModule.zip"
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
if (Test-Path "$targetondisk\EWSModule") { Remove-Item -Force "$targetondisk\EWSModule" }
Rename-Item -Path ($targetondisk+"\EWSModule-master") -NewName "EWSModule" -Force
Write-Host "Module has been installed" -ForegroundColor Green
Import-Module -Name EWSModule
Get-Command -Module EWSModule