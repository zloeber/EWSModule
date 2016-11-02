function Get-EWSRootFolderID {
    <#
    .SYNOPSIS
        Return a mailbox folder object.
    .DESCRIPTION
        Return a mailbox folder object.
    .PARAMETER EWSService
        Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
        Mailbox to target. If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.
    .PARAMETER FolderBase
        A well known folder base name (Inbox, Calendar, Contacts, et cetera..)

    .EXAMPLE
        PS > Get-EWSRootFolderID -EWSService $EWSService -FolderRoot Contacts -Mailbox 'jdoe@contoso.com'

        Description
        -----------
        Return the Folder object for the currently connected EWSService account of the well known 'contacts' folder
        ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::contacts)

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox of folder.')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Well known folder object.')]
        [ValidateSet('Calendar','Contacts','DeletedItems','Drafts','Inbox','Journal','Notes','Outbox','SentItems','Tasks','MsgFolderRoot','PublicFoldersRoot','Root','JunkEmail','SearchFolders','VoiceMail','RecoverableItemsRoot','RecoverableItemsDeletions','RecoverableItemsVersions','RecoverableItemsPurges','ArchiveRoot','ArchiveMsgFolderRoot','ArchiveDeletedItems','ArchiveRecoverableItemsRoot','ArchiveRecoverableItemsDeletions','ArchiveRecoverableItemsVersions','ArchiveRecoverableItemsPurges','SyncIssues','Conflicts','LocalFailures','ServerFailures','RecipientCache','QuickContacts','ConversationHistory','ToDoSearch')]
        [string]$FolderBase = 'MsgFolderRoot'
    )
    # Pull in all the caller verbose,debug,info,warn and other preferences
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    
    if ($EWSService -eq $null) {
        Write-Verbose "$($FunctionName): Using module local ews service object"
        $EWSService = Get-EWSService
    }
    
    if ($EWSService -eq $null) {
        throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
    }

    try {
        $TargetedMailbox = Get-EWSTargetedMailbox -EWSService $EWSService -Mailbox $Mailbox
        Write-Verbose "$($FunctionName): Targeted Mailbox = $($TargetedMailbox) "
        $ConnectedMB = New-Object ews_mailbox($TargetedMailbox)
        new-object ews_folderid([ews_wellknownfolder]::$FolderBase,$ConnectedMB)
    }
    catch {
        throw "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
    }
}