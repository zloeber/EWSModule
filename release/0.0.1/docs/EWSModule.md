---
Module Name: EWSModule
Module Guid: 00000000-0000-0000-0000-000000000000
Download Help Link: https://github.com/zloeber/EWSModule/release/EWSModule/docs/EWSModule.md
Help Version: 0.0.2
Locale: en-US
---

# EWSModule Module
## Description
An easier way to use EWS with Powershell

## EWSModule Cmdlets
### [Connect-EWS](Connect-EWS.md)
Connects to Exchange Web Services.

### [Get-EmailAddressFromAD](Get-EmailAddressFromAD.md)
Return the email address of a User ID from AD.

### [Get-EWSCalendarAppointments](Get-EWSCalendarAppointments.md)
Uses the much faster FindItems as opposed to FindAppointments to return calendar appointments.

### [Get-EWSCalenderViewAppointments](Get-EWSCalenderViewAppointments.md)
Uses a slower method for accessing and returning calendar appointments

### [Get-EWSContact](Get-EWSContact.md)
Gets a single contact in a Contact folder within a mailbox using the Exchange Web Services API

### [Get-EWSContacts](Get-EWSContacts.md)
Gets all contacts in a Contact folder in a Mailbox using the Exchange Web Services API

### [Get-EWSDllLoadState](Get-EWSDllLoadState.md)
Determine if the EWS dll is loaded or not.

### [Get-EWSFolder](Get-EWSFolder.md)
Return a mailbox folder object.

### [Get-EWSFolderItem](Get-EWSFolderItem.md)
Returns items from a mailbox folder.

### [Get-EWSFolderPaths](Get-EWSFolderPaths.md)
Return a mailbox folder object.

### [Get-EWSModuleInitializationState](Get-EWSModuleInitializationState.md)
Returns the initialization state of the module.

### [Get-EWSOOFSettings](Get-EWSOOFSettings.md)
Get the out of office settings for a mailbox.

### [Get-EWSRootFolderID](Get-EWSRootFolderID.md)
Return a mailbox folder object.

### [Get-EWSService](Get-EWSService.md)
Returns the current EWSService module variable object

### [Get-EWSTargetedMailbox](Get-EWSTargetedMailbox.md)
Return the intended targeted email address of a mailbox for ews operations.

### [Get-ServerCertificateValidationCallback](Get-ServerCertificateValidationCallback.md)
Returns the current certificate validation callback setting

### [Import-EWSDll](Import-EWSDll.md)
Load EWS dlls.

### [Initialize-EWS](Initialize-EWS.md)
Load EWS dlls and create type accelerators for other functions.

### [Install-EWSDll](Install-EWSDll.md)
Attempts to download and extract the ews dll needed for this library.

### [New-EWSCalendarEntry](New-EWSCalendarEntry.md)
Creates an appointment object that can be manipulated or saved.

### [New-EWSExtendedProperty](New-EWSExtendedProperty.md)
Creates a new extended property which can be assigned to items in outlook.

### [Remove-EWSCalendarAppointment](Remove-EWSCalendarAppointment.md)
Remove a calendar appointment object from a mailbox.

### [Set-EWSMailboxImpersonation](Set-EWSMailboxImpersonation.md)
Set the impersonation for a mailbox.

### [Set-EWSModuleInitializationState](Set-EWSModuleInitializationState.md)
Sets the module initialization state

### [Set-EWSOofSettings](Set-EWSOofSettings.md)
Set the out of office settings for a mailbox.

### [Set-EWSService](Set-EWSService.md)
Sets the module connected service.

### [Set-EWSSSLIgnoreWorkaround](Set-EWSSSLIgnoreWorkaround.md)
Sets the module to ignore SSL checking.

### [Set-ServerCertificateValidationCallback](Set-ServerCertificateValidationCallback.md)
Sets the current certificate validation callback setting

### [Test-EWSAutodiscover](Test-EWSAutodiscover.md)
This function uses the EWS Managed API to test the Exchange Autodiscover service.



