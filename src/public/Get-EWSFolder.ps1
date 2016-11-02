function Get-EWSFolder {
    <#
    .SYNOPSIS
    Return a mailbox folder object.
    .DESCRIPTION
    Return a mailbox folder object.
    .PARAMETER EWSService
    Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
    Mailbox to target. If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.
    .PARAMETER FolderPath
    Path of folder in the form of /folder1/folder2

    .EXAMPLE
    PS > Get-EWSFolder -FolderRoot Contacts

    Return the Folder object for the currently connected EWSService account of the well known 'contacts' folder
    ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::contacts)
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
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox of folder.')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Folder path.')]
        [string]$FolderPath = '\',
        [parameter(Position=3, HelpMessage='Well known folder object.')]
        [ValidateSet('Calendar','Contacts','DeletedItems','Drafts','Inbox','Journal','Notes','Outbox','SentItems','Tasks','MsgFolderRoot','PublicFoldersRoot','Root','JunkEmail','SearchFolders','VoiceMail','RecoverableItemsRoot','RecoverableItemsDeletions','RecoverableItemsVersions','RecoverableItemsPurges','ArchiveRoot','ArchiveMsgFolderRoot','ArchiveDeletedItems','ArchiveRecoverableItemsRoot','ArchiveRecoverableItemsDeletions','ArchiveRecoverableItemsVersions','ArchiveRecoverableItemsPurges','SyncIssues','Conflicts','LocalFailures','ServerFailures','RecipientCache','QuickContacts','ConversationHistory','ToDoSearch')]
        [string]$SearchBase = 'MsgFolderRoot'
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
    }
    catch {
        throw "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
    }
    
    try {
        Write-Verbose "$($FunctionName): Trying to connect to get the EWS folder ID of $($SearchBase)"
        $FolderRootID = Get-EWSRootFolderID -EWSService $EWSService -FolderBase $SearchBase -Mailbox $TargetedMailbox
        $FolderRoot = [ews_folder]::Bind($EWSService,$FolderRootID)
        $Folder = $FolderRoot
    }
    catch {
        throw "$($FunctionName): $($_)"
    } 

    # If we have multiple paths deep then we need to go through each of them one after another and search for the folder
    if ($FolderPath -ne '\') {
        # Split our folder path to individual folder elements
        $PathElements = $FolderPath -split '\\' 
        For ($i=0; $i -lt $PathElements.Count; $i++) {
            if ($PathElements[$i]) {
                Write-Verbose  "$($FunctionName): Processing path - $($PathElements[$i])"
                $View = New-Object ews_folderview(2,0) 
                $View.PropertySet = [ews_basepropset]::IdOnly
                $SearchFilter = New-Object ews_searchfilter_isequalto([ews_schema_folder]::DisplayName, $PathElements[$i])
                $FolderResults = $Folder.FindFolders($SearchFilter, $View) 
                if ($FolderResults.TotalCount -ne 1) { 
                    # We have either none or more than one folder returned... Either way, we can't continue 
                    Write-Error "$($FunctionName): Failed to find $($PathElements[$i]), path requested was $FolderPath"
                }
                    
                # Connect to the folder for the next search
                $Folder = [ews_folder]::Bind($EWSService,$FolderResults.Folders[0].Id)           
            }
        } 
    }
    return $Folder
}