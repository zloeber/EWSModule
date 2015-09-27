function Get-EWSFolderPaths {
    <#
    .SYNOPSIS
        Return a mailbox folder object.
    .DESCRIPTION
        Return a mailbox folder object.
    .PARAMETER EWSService
        Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER RootFolderId
        Folder to target. Can target specific mailboxes with Get-EWSFolder
    .PARAMETER FolderCache
        
    .PARAMETER FolderPrefix
        

    .EXAMPLE
        PS > 
        PS > 

        Description
        -----------
        TBD

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [Parameter(Mandatory=$true)] 
        [Microsoft.Exchange.WebServices.Data.FolderId]$RootFolderId,
        [Parameter(Mandatory=$true)]
        [PSObject]$FolderCache,
        [Parameter(Position=3)]
        [String]$FolderPrefix
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
    
    #Define Extended properties  
    $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)  
    $PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long)
    
    #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
    $fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
    
    #Deep transversal will ensure all folders in the search path are returned  
    $fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
    
    #Add Properties to the  Property Set  
    $psPropertySet.Add($PR_Folder_Path)
    $psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED)
    $fvFolderView.PropertySet = $psPropertySet 
    
    #The Search filter will exclude any Search Folders  
    $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")
    $fiResult = $null
    
    #The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
    do {  
        $fiResult = $EWSService.FindFolders($RootFolderId,$sfSearchFilter,$fvFolderView)  
        foreach($ffFolder in $fiResult.Folders) {
            #Try to get the FolderPath Value and then covert it to a usable String 
            $foldpathval = $null
            if ($ffFolder.TryGetProperty($PR_Folder_Path,[ref]$foldpathval)) {  
                $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
                $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
                $hexString = $hexArr -join ''  
                $hexString = $hexString.Replace("FEFF", "5C00")  
                $fpath = ConvertTo-String($hexString)  
            }
            if($FolderCache.ContainsKey($ffFolder.Id.UniqueId) -eq $false) {
                if ([string]::IsNullOrEmpty($FolderPrefix)) {
                    $FolderCache.Add($ffFolder.Id.UniqueId,($fpath))    
                }
                else {
                    $FolderCache.Add($ffFolder.Id.UniqueId,("\" + $FolderPrefix + $fpath))    
                }
            }
        } 
        $fvFolderView.Offset += $fiResult.Folders.Count
    } while($fiResult.MoreAvailable)
}