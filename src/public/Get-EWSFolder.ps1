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
    .PARAMETER PublicFolder
        Force target a public folder instead.

    .EXAMPLE
        PS > Get-EWSFolder -FolderObject Contacts

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
    [CmdletBinding(DefaultParametersetName='FolderAsString')]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [parameter(ParameterSetName='FolderAsString')]
        [parameter(ParameterSetName='FolderAsObject')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox of folder.')]
        [parameter(ParameterSetName='FolderAsString')]
        [parameter(ParameterSetName='FolderAsObject')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Folder path.')]
        [parameter(ParameterSetName='FolderAsString')]
        [string]$FolderPath,
        [parameter(Position=2, HelpMessage='Well known folder object.')]
        [parameter(ParameterSetName='FolderAsObject')]
        [ValidateSet('Calendar','Contacts','DeletedItems','Drafts','Inbox','Journal','Notes','Outbox','SentItems','Tasks','MsgFolderRoot','PublicFoldersRoot','Root','JunkEmail','SearchFolders','VoiceMail','RecoverableItemsRoot','RecoverableItemsDeletions','RecoverableItemsVersions','RecoverableItemsPurges','ArchiveRoot','ArchiveMsgFolderRoot','ArchiveDeletedItems','ArchiveRecoverableItemsRoot','ArchiveRecoverableItemsDeletions','ArchiveRecoverableItemsVersions','ArchiveRecoverableItemsPurges','SyncIssues','Conflicts','LocalFailures','ServerFailures','RecipientCache','QuickContacts','ConversationHistory','ToDoSearch')]
        [ews_wellknownfolder]$FolderObject,
        [parameter(Position=3, HelpMessage='Are you targeting a public Folder Path?')]
        [parameter(ParameterSetName='FolderAsString')]
        [switch]$PublicFolder
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
    
    # Return a reference to a folder specified by path 
    
    switch ($PsCmdlet.ParameterSetName) { 
        'FolderAsString' {
            if ($PublicFolder) { 
                $mbx = ''
                try {
                    $Folder = [ews_folder]::Bind($EWSService, [ews_wellknownfolder]::PublicFoldersRoot) 
                }
                catch {
                    Write-Warning "$($FunctionName): Unable to find a public folder server or database to connect to."
                    return $null
                }
            }
            else {
                $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
                $mbx = New-Object ews_mailbox($email)
                $FolderID = New-Object ews_folderid([ews_wellknownfolder]::MsgFolderRoot, $mbx )
            }
            
            if ($FolderPath -ne '\') {
                $PathElements = $FolderPath -split '\\' 
                For ($i=0; $i -lt $PathElements.Count; $i++) { 
                    if ($PathElements[$i]) { 
                        $View = New-Object ews_folderview(2,0) 
                        $View.PropertySet = [ews_basepropset]::IdOnly
                        $SearchFilter = New-Object ews_searchfilter_isequalto([ews_schema_folder]::DisplayName, $PathElements[$i])
                        $FolderResults = $Folder.FindFolders($SearchFilter, $View) 
                        if ($FolderResults.TotalCount -ne 1) { 
                            # We have either none or more than one folder returned... Either way, we can't continue 
                            Write-Verbose "$($FunctionName): Failed to find $($PathElements[$i]), path requested was $FolderPath"
                            return $null
                        }
                         
                        if (-not [String]::IsNullOrEmpty(($mbx))) {
                            $folderId = New-Object ews_folderid($FolderResults.Folders[0].Id, $mbx) 
                            try {
                                $Folder = [ews_folder]::Bind($service, $folderId) 
                            }
                            catch {
                                Write-Warning "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
                                return $null
                            }
                        } 
                        else {
                            try {
                                $Folder = [ews_folder]::Bind($service, $FolderResults.Folders[0].Id)
                            }
                            catch {
                                Write-Warning "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
                                return $null
                            }
                        } 
                    } 
                } 
            }
            else {
                try {
                    $Folder = [ews_folder]::Bind($EWSService, $FolderID)
                }
                catch {
                    Write-Warning "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
                    return $null
                }
            }
        }
        'FolderAsObject' {
            $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
            $mbx = New-Object ews_mailbox($email)
            $FolderID = New-Object ews_folderid($FolderObject, $mbx)
            try {
                $Folder = [ews_folder]::Bind($EWSService, $FolderID)
            }
            catch {
                Write-Warning "$($FunctionName): Unable to connect to the specified folder. Check that you have permissions to access this mailbox"
                return $null
            }
        }
    }
    return $Folder 
}