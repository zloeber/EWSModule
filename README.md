# EWSModule - A stupid easy Powershell EWS module

.... or at least my attempt at one anyway. More documentation to come but if you want to get started download the module anywhere and run some code like the following:
 
 Import-Module ./EWSModule.psm1
 
 Install-EWSDLL -Verbose
 
 Initialize-EWS -Verbose
 
 Connect-EWS -Credential (Get-Credential) -Verbose
 
 Get-EWSFolder
