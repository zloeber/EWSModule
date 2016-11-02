function Get-EWSFolderItem {
    <#
    .SYNOPSIS
    Returns items from a mailbox folder.
    .DESCRIPTION
    Returns items from a mailbox folder using either an AQS search filter or just by a number of desired results.
    .PARAMETER EWSService
    Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
    Mailbox to target. If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.
    .PARAMETER FolderPath
    A specific path to search.
    .PARAMETER SearchBase
    Base folder to return items from. The default is Inbox.
    .PARAMETER Count
    Number of items to return if not using a search string. With a search string it is the number of items returned per page
    .PARAMETER Filter
    Search string in AQS format for returning items. For complete documentation on syntax see https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx
    .EXAMPLE
    PS > Get-EWSFolderItem -Count 10 -SearchString 'Subject:Change Notice' -verbose

    Retrieves all emails containing the term 'Change Notice' in the subject from the Inbox. Results are fetched 10 at a time and verbose
    output is displayed on the screen.
    .EXAMPLE
    PS > Get-EWSFolderItem -Count 10 -SearchString 'Received:Yesterday' -verbose

    Retrieves all emails in the connected user's Inbox received yesterday. Results are fetched 10 at a time and verbose
    output is displayed on the screen.    
    .LINK
    http://www.the-little-things.net/
    .LINK
    https://www.github.com/zloeber/EWSModule
    .LINK
    https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx
    .NOTES
    Author: Zachary Loeber
    Requires: Powershell 3.0
    Version History
    1.0.0 - Initial release
    #>
    [CmdletBinding(DefaultParameterSetName='Count')]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox of folder.')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Folder path.')]
        [string]$FolderPath = '\',
        [parameter(Position=3, HelpMessage='Well known folder object.')]
        [ValidateSet('Calendar','Contacts','DeletedItems','Drafts','Inbox','Journal','Notes','Outbox','SentItems','Tasks','MsgFolderRoot','PublicFoldersRoot','Root','JunkEmail','SearchFolders','VoiceMail','RecoverableItemsRoot','RecoverableItemsDeletions','RecoverableItemsVersions','RecoverableItemsPurges','ArchiveRoot','ArchiveMsgFolderRoot','ArchiveDeletedItems','ArchiveRecoverableItemsRoot','ArchiveRecoverableItemsDeletions','ArchiveRecoverableItemsVersions','ArchiveRecoverableItemsPurges','SyncIssues','Conflicts','LocalFailures','ServerFailures','RecipientCache','QuickContacts','ConversationHistory','ToDoSearch')]
        [string]$SearchBase = 'Inbox',
        [parameter(Position=4, HelpMessage='Number of items to return if not using a search string. With a search string it is the number of items returned per page.')]
        [int]$Count = 1,
        [parameter(Position=5, HelpMessage='Search string in AQS format.')]
        [string]$Filter  
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
    }
    catch {
        throw "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
    }

    try {
        $Folder = Get-EWSFolder -EWSService $EWSService -Mailbox $TargetedMailbox -FolderPath $FolderPath -SearchBase $SearchBase
    }
    catch {
        Write-Error "$($FunctionName): Unable to locate the folder to search"
    }

    if ([string]::IsNullOrEmpty($Filter)) {
        Write-Verbose "$($FunctionName): No AQS query sent. Returning the top $($Count) items."
        $Items = @($Folder.FindItems($Count))
        Write-Verbose "$($FunctionName): Found $($Items.Count) items."
        # Load additional item content and return the found item.
        if ($Items.Count -gt 0) {
            $Items | Foreach {
                $_.Load()
                $_
            }
        }
    }
    else {
        Write-Verbose "$($FunctionName): AQS query sent."
        #Define the properties to get
        $psPropset = new-object ews_propset([ews_basepropset]::FirstClassProperties)    
        $ivItemView =  New-Object ews_itemview($Count)
        $Items = $null
        $Page = 1
        do{
            $Items = $Folder.FindItems($Filter, $ivItemView)
            # Get the extra properties for all found items
            if ($Items.Count -gt 0) {
                Write-Verbose "$($FunctionName): Processing batch number $($Page) of found items."
                $null = $EWSService.LoadPropertiesForItems($Items,$psPropset)
                # Return the results of this batch
                $Items
                $ivItemView.Offset += $Items.Items.Count
                $Page++
            }
            else {
                Write-Verbose "$($FunctionName): No results found for this search."
            }
        } While ($Items.MoreAvailable -eq $true)
    }
}