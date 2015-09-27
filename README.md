# EWSModule - A stupid easy Powershell EWS module

##Description
.... or at least my attempt at one anyway. More documentation to come but if you want to get started download the module  anywhere and run some code like the following:

```
Import-Module ./EWSModule.psm1
Install-EWSDLL -Verbose
Initialize-EWS -Verbose
Connect-EWS -Credential (Get-Credential) -Verbose
Get-EWSFolder
```

##Install

You can install this module manually with the following command.

`iex (New-Object Net.WebClient).DownloadString("https://github.com/zloeber/EWSModule/raw/master/Install-EWSModule.ps1")`

Note that this module does not come with the prerequisite EWS dll (primarily because I simply don't know the rules around redistributing that dll). You can still load the module up then install the DLL with the Install-EWSDLL module (assuming that Initialize-EWS fails that is, if Initialize-EWS works then you likely already have the EWS dlls installed to some default location and are set to go).

##Status
The initial code behind this was for a project I was working on based around calendar appointments. This code was based on an even earlier one off script I wrote just to get a job done. As I got some internal requests for the code I was able to refine some of the functions and used what was created as an opportunity to become better with PowerShell modules. So, pull requests and improvement suggestions are certainly welcome.

That being said, I've only tested in production a handful of the functions. You should read that as a warning to test thoroughly anything you might construct with this module in a test environment first as I make zero guarantees on the fidelity of all this code.

A large number of the exported cmdlets are not even finished yet! I've started dissecting [Glen Scale's work](http://gsexdev.blogspot.com/) and integrating it into this module but haven't made a large amount of progress on it yet. This includes all the contact cmdlets and some of the private functions as well. (Glen, if you are reading this, your work is brilliant but don’t take offense if I take huge liberties in cleaning up your code).

##Exported Cmdlets
Here is a general list of functions included with this module.
<table>
<tr><th>Name</th><th>Status</th></tr>
<tr><td>Connect-EWS</td><td>Complete</td></tr>
<tr><td>Copy-EWSContactsFromGalToMailbox</td><td>In Progress</td></tr>
<tr><td>Create-EWSContact</td><td>In Progress</td></tr>
<tr><td>Create-EWSContactGroup</td><td>In Progress</td></tr>
<tr><td>Delete-EWSContact</td><td>In Progress</td></tr>
<tr><td>Export-EWSContact</td><td>In Progress</td></tr>
<tr><td>Export-EWSGALContact</td><td>In Progress</td></tr>
<tr><td>Get-EmailAddressFromAD</td><td>Complete</td></tr>
<tr><td>Get-EWSAttachmentTypeMailboxStats</td><td>In Progress</td></tr>
<tr><td>Get-EWSAutoDiscoverPhotoURL</td><td>In Progress</td></tr>
<tr><td>Get-EWSCalendarAppointments</td><td>Complete</td></tr>
<tr><td>Get-EWSCalenderViewAppointments</td><td>Complete</td></tr>
<tr><td>Get-EWSContact</td><td>In Progress</td></tr>
<tr><td>Get-EWSContactFolder</td><td>In Progress</td></tr>
<tr><td>Get-EWSContactGroup</td><td>In Progress</td></tr>
<tr><td>Get-EWSContacts</td><td>In Progress</td></tr>
<tr><td>Get-EWSDllLoadState</td><td>Complete</td></tr>
<tr><td>Get-EWSeDiscoveryKeyWordStats</td><td>In Progress</td></tr>
<tr><td>Get-EWSFolder</td><td>Complete</td></tr>
<tr><td>Get-EWSFolderPaths</td><td>In Progress</td></tr>
<tr><td>Get-EWSMailboxAttachments</td><td>In Progress</td></tr>
<tr><td>Get-EWSMailboxConversationStats</td><td>Complete</td></tr>
<tr><td>Get-EWSMailboxItemStats</td><td>In Progress</td></tr>
<tr><td>Get-EWSMailboxItemTypeStats</td><td>In Progress</td></tr>
<tr><td>Get-EWSModuleInitializationState</td><td>Complete</td></tr>
<tr><td>Get-EWSOofSettings</td><td>Complete</td></tr>
<tr><td>Get-EWSService</td><td>Complete</td></tr>
<tr><td>Get-ServerCertificateValidationCallback</td><td>Complete</td></tr>
<tr><td>Import-EWSDll</td><td>Complete</td></tr>
<tr><td>Initialize-EWS</td><td>Complete</td></tr>
<tr><td>Install-EWSDll</td><td>Complete</td></tr>
<tr><td>New-EWSCalendarEntry</td><td>Complete</td></tr>
<tr><td>New-EWSeDiscoveryPreviewItems</td><td>In Progress</td></tr>
<tr><td>New-EWSeDiscoveryPreviewItemsStats</td><td>In Progress</td></tr>
<tr><td>New-EWSExtendedProperty</td><td>Complete</td></tr>
<tr><td>Set-EWSMailboxImpersonation</td><td>Complete</td></tr>
<tr><td>Set-EWSModuleInitializationState</td><td>Complete</td></tr>
<tr><td>Set-EWSOofSettings</td><td>Complete</td></tr>
<tr><td>Set-EWSService</td><td>Complete</td></tr>
<tr><td>Set-EWSSSLIgnoreWorkaround</td><td>Complete</td></tr>
<tr><td>Set-ServerCertificateValidationCallback</td><td>Complete</td></tr>
<tr><td>Test-EWSAutodiscover</td><td>Complete</td></tr>
<tr><td>Update-EWSContact</td><td>In Progress</td></tr>
</table>


##Credits
Glen Scale - [Blog](http://gsexdev.blogspot.com/), [Github](https://github.com/gscales)

##Other Information
**Author:** Zachary Loeber

**Website:** [http://www.the-little-things.net](http://www.the-little-things.net)

**Github:** [https:/github.com/zloeber/EWSModule](https:/github.com/zloeber/EWSModule)
