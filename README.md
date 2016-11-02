# EWSModule - A stupid easy Powershell EWS module

##Description
.... or at least my attempt at one anyway. If you want to get started download the module anywhere and run code like this to get the a list of the current user's mailbox folders:

```
Import-Module ./EWSModule.psm1
Connect-EWS -Credential (Get-Credential) -Verbose
Get-EWSFolderItem -SearchBase Inbox -Count 5
```

##Install

You can install this module manually with the following command.

`iex (New-Object Net.WebClient).DownloadString("https://github.com/zloeber/EWSModule/raw/master/Install-EWSModule.ps1")`

~~Note that this module does not come with the prerequisite EWS dll (primarily because I simply don't know the rules around redistributing that dll). You can still load the module up then install the DLL with the Install-EWSDLL module (assuming that Initialize-EWS fails that is, if Initialize-EWS works then you likely already have the EWS dlls installed to some default location and are set to go).~~

##Status
The initial code behind this was for a project I was working on based around calendar appointments. This code was based on an even earlier script I wrote just to get a job done. As I got some internal requests for the code, I refined some of the functions and used this as an opportunity to become better with PowerShell modules. So, pull requests and improvement suggestions are certainly welcome.

That being said, I've only tested in production a handful of the functions. You should read that as a warning to test thoroughly anything you might construct with this module in a test environment first as I make zero guarantees on the fidelity of all this code.

A large number of the exported cmdlets are not even finished yet! I've started dissecting [Glen Scale's work](http://gsexdev.blogspot.com/) and integrating it into this module. This includes all the contact cmdlets and some of the private functions as well. (Glen, if you are reading this, your work is brilliant but don’t take offense if I take huge liberties in cleaning it up).

##Module Notes
I made some coding decisions I'm not entirely proud of as well as some design choices worth documenting. So I felt they should be mentioned here along with any other interesting facts one might want to know.

###Module Initialization
I didn't want to have to worry about EWS DLL redistribution licensing or keeping it updated should Microsoft ever decide to release a new version. This means that one cannot simply load the module and be ready to rock and roll. We need to factor in locating and loading the appropriate DLL. I know we can pass parameters to modules when we load them but I felt that this would be counter-intuitive (do you know of any modules you use regularly that you have to do this with to get them to load?). So I opted to include an initialization function which needs to be called after the module is loaded. This function can be passed the location of the EWS DLL via the  `-EWSManagedApiPath` flag. Otherwise it tries a handful of default locations (including the current directory).

If you don't want to install the EWS managed API then all we really need is a single DLL. I've made it easy to automatically download and extract the appropriate DLL from the MSI with a single command:

`Install-EWSDLL`

You don't even have to install this module to use it honestly. You can clone the repo, load the module from the downloaded folder, download and use the most recent EWS dll, initialize, and then connect to EWS in just a few simple commands.

**Update:** I've updated the code to be even more simple to use. If you use connect-ews with a valid credential a number of locations will be checked for the dll. If it isn't found it will be automatically extracted from a base64 encoded variable into the module directory for use instead. So all you need to do is use connect-ews to; install the appropriate dll, run the import of the dll into memory, initialize all the type accelerators, and connect to EWS.

###Custom Type Accelerators
Anyone who has worked at all with Exchange Web Services is aware of the large number of .NET type references which are required to get anything done. This tends to make any code horribly complex and scary looking. Go ahead and [take a look at some right now](https://raw.githubusercontent.com/gscales/Powershell-Scripts/master/EWSContacts/EWSContactFunctions.ps1) if you want to be overwhelmed.

So early on, as a bit of a learning exercise, I included an initialization routine to add a bunch of custom type accelerators to shorten up the code a bit. An example would be turning this beast:

`[Microsoft.Exchange.WebServices.Data.BasePropertySet]`

into just:

`[ews_basepropset]`

I went a bit nuts with these and just decided to roll with it in this module across all functions. This means you will see all sorts of unfamiliar type references in the code. These custom type accelerators get added as part of the Initialize-EWS function call that is required to further use the module functions (this same function is called with an -Uninitialize flag to remove any custom type accelerator with ews_ in the name when the module is unloaded).

Interestingly enough I'm able to load up and export functions that have these custom type accelerators in strongly typed parameters before they are ever defined. It seems there is something worth researching here on the inner workings of PowerShell when I get some free time.

Anyway, if you are going through the code you can find the .NET type to accelerator name lookup in a hash called `$EWSAccels` within the base EWSModule.psm1 file. 

###Caller Preferences
I've personally run into issues when working with larger PowerShell projects where I'm using a custom advanced function within another advanced function and it fails to adhere to the -Verbose flag I'm sending to the parent function. In the past I came up with all kinds of silly work arounds for this. Now I'm using another author's clever function, Get-CallerPreference, in every function I create. This should mean that if you are using this module to build larger projects you can confidently know that your preferences will be followed down the stack (unless explicitly set obviously).

###EWS Connections
This module is designed to reuse a single EWS connection across all functions. There is a private variable that stores the results of Connect-EWS for use in all the other functions. This should speed up script processing a bit and reduce resource utilization. You can also store the connection using `Get-EWSConnection` into a variable and pass it manually along to most of the cmdlets if you need to do so.

###Targeting Mailboxes
I've removed all `Impersonate` flags from any imported cmdlets as you can infer impersonation is occuring by the connection state of EWS. If you want to impersonate a user you will need to use the `Set-EWSMailboxImpersonation` cmdlet. The rest of the cmdlets will then use `Get-EWSTargettedMailbox` to figure out which mailbox an operation should target.

Mailbox targeting is done in this order:
1. If a mailbox name is passed to the function and it is a valid email address we will lookup and try to target it directly without impersonation.
2. If a mailbox name is passed to the function and it is NOT a valid email address we will try to lookup the email address for the assumed account ID in Active Directory and try to target it directly without impersonation.
2. If no mailbox name is passed to the function then we will look for the ImpersonationUserID of the EWS connection. If it is set we will attempt to target that mailbox via impersonation.
3. If ImpersonationUserID is not set then we will try and target the EWS credential UserID directly. If it is not in email address format we attempt to resolve it to the correct format using EWS directory lookup mechanisms (not a direct AD query).

### Building
If you want to build this module on your own I've included the entire build script needed to automate the process. One note is that the script analyzer will run and find that we are passing direct username/password parameters in 'Connect-EWS' (it also takes the appropriate credential object for what it is worth). This will normally error out the build process but I've put in a yes/no prompt to continue anyway. To build the script simply download this repo and run:

```
.\build.ps1
```

##Credits
Glen Scale - [Blog](http://gsexdev.blogspot.com/), [Github](https://github.com/gscales)
David Wyatt  - [Get-CallerPreference](https://gallery.technet.microsoft.com/scriptcenter/Inherit-Preference-82343b9d)

##Other Information
**Author:** Zachary Loeber

**Website:** [http://www.the-little-things.net](http://www.the-little-things.net)

**Github:** [https:/github.com/zloeber/EWSModule](https:/github.com/zloeber/EWSModule)
