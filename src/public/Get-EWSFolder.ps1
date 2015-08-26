function Get-EWSFolder {
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(Position=1, HelpMessage='Mailbox of folder.')]
        [string]$Mailbox,
        [parameter(Position=2, HelpMessage='Folder path.')]
        [string]$FolderPath,
        [parameter(Position=2, HelpMessage='Public Folder Path?')]
        [switch]$PublicFolder
    )
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    
    if ($EWSService -eq $null) {
        Write-Verbose "$($FunctionName): Using module local ews service object"
        $EWSService = Get-EWSService
        Write-Verbose "$($FunctionName): URL targeted = $($EWSService.URL)"
    }
    
    if ($EWSService -eq $null) {
        throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
    }
    
    # Return a reference to a folder specified by path 
    if ($PublicFolders) { 
        $mbx = '' 
        $Folder = [ews_folder]::Bind($EWSService, [ews_wellknownfolder]::PublicFoldersRoot) 
    } 
    else {
        $email = Get-EWSTargettedMailbox -EWSService $EWSService -Mailbox $Mailbox
        $mbx = New-Object ews_mailbox( $email ) 
        $folderId = New-Object ews_folderid([ews_wellknownfolder]::MsgFolderRoot, $mbx ) 
        $Folder = [ews_folder]::Bind($EWSService, $folderId) 
    } 
 
    if ($FolderPath -ne '\') {
        $PathElements = $FolderPath -split '\\' 
        For ($i=0; $i -lt $PathElements.Count; $i++) { 
            if ($PathElements[$i]) { 
                $View = New-Object  ews_folderview(2,0) 
                $View.PropertySet = [ews_basepropset]::IdOnly
                $SearchFilter = New-Object ews_searchfilter_isequalto([ews_schema_folder]::DisplayName, $PathElements[$i])
                $FolderResults = $Folder.FindFolders($SearchFilter, $View) 
                if ($FolderResults.TotalCount -ne 1) { 
                    # We have either none or more than one folder returned... Either way, we can't continue 
                    $Folder = $null 
                    Write-Verbose "$($FunctionName): Failed to find $($PathElements[$i]), path requested was $FolderPath"
                    break 
                }
                 
                if (-not [String]::IsNullOrEmpty(($mbx))) {
                    $folderId = New-Object ews_folderid($FolderResults.Folders[0].Id, $mbx ) 
                    $Folder = [ews_folder]::Bind($service, $folderId) 
                } 
                else {
                    $Folder = [ews_folder]::Bind($service, $FolderResults.Folders[0].Id) 
                } 
            } 
        } 
    }

    return $Folder 
}