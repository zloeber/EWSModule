function New-EWSeDiscoveryPreviewItemsStats {
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [Parameter(Position=1, Mandatory=$true)] [String]$KQL,
        [Parameter(Position=2, Mandatory=$true)] [String]$SearchableMailboxString
        
    )  
    $FolderCache = @{}
    # Bind to the MsgFolderRoot folder  
    $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)   
    $MsgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
    Get-EWSFolderPaths -FolderCache $FolderCache -service $service -rootFolderId $MsgRoot.Id
    try
    {
        # Bind to the ArchiveMsgFolderRoot folder  
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$MailboxName)   
        $ArchiveRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
        Get-EWSFolderPaths -FolderCache $FolderCache -service $service -rootFolderId $ArchiveRoot.Id -FolderPrefix "Archive"
        Write-Host ("Mailbox has Archive")
    }
    catch
    {
    
    }
    $gsMBResponse = $service.GetSearchableMailboxes($SearchableMailboxString, $false);
    $msbScope = New-Object  Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $gsMBResponse.SearchableMailboxes.Length
    $mbCount = 0;
    foreach ($sbMailbox in $gsMBResponse.SearchableMailboxes)
    {
        $msbScope[$mbCount] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($sbMailbox.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All);
        $mbCount++;
    }
    $smSearchMailbox = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
    $mbq =  New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $msbScope);
    $mbqa = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
    $mbqa[0] = $mbq
    $smSearchMailbox.SearchQueries = $mbqa;
    $smSearchMailbox.PageSize = 1000;
    $smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
    $smSearchMailbox.PerformDeduplication = $false;           
    $smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::PreviewOnly;
    $srCol = $service.SearchMailboxes($smSearchMailbox);
    $rptCollection = @{}

    if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
    {
        Write-Host ("Items Found " + $srCol[0].SearchResult.ItemCount)
        if ($srCol[0].SearchResult.ItemCount -gt 0)
        {                  
            do
            {
                $smSearchMailbox.PageItemReference = $srCol[0].SearchResult.PreviewItems[$srCol[0].SearchResult.PreviewItems.Length - 1].SortValue;
                foreach ($PvItem in $srCol[0].SearchResult.PreviewItems) {
                    $rptObj = "" | select FolderPath,TotalItemNumbers,Size
                    if($FolderCache.ContainsKey($PvItem.ParentId.UniqueId)){
                        if($rptCollection.ContainsKey($PvItem.ParentId.UniqueId) -eq $false){
                            $rptObj = "" | Select FolderPath,TotalItemNumbers,TotalSize
                            $rptObj.FolderPath = $FolderCache[$PvItem.ParentId.UniqueId]
                            $rptCollection.Add($PvItem.ParentId.UniqueId,$rptObj)
                        }
                        $rptCollection[$PvItem.ParentId.UniqueId].TotalSize += $PvItem.Size
                        $rptCollection[$PvItem.ParentId.UniqueId].TotalItemNumbers++
                    }
                    else{
                    #$ItemId = new-object Microsoft.Exchange.WebServices.Data.ItemId($PvItem.Id)   
                        $Item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$PvItem.Id)
                        if($FolderCache.ContainsKey($Item.ParentFolderId.UniqueId)){
                            $FolderCache.Add($PvItem.ParentId.UniqueId,$FolderCache[$Item.ParentFolderId.UniqueId])
                            if($rptCollection.ContainsKey($PvItem.ParentId.UniqueId) -eq $false){
                                $rptObj = "" | Select FolderPath,TotalItemNumbers,TotalSize
                                $rptObj.FolderPath = $FolderCache[$PvItem.ParentId.UniqueId]
                                $rptCollection.Add($PvItem.ParentId.UniqueId,$rptObj)
                            }
                            $rptCollection[$PvItem.ParentId.UniqueId].TotalSize += $PvItem.Size
                            $rptCollection[$PvItem.ParentId.UniqueId].TotalItemNumbers++
                        }
                    }
                }                        
                $srCol = $service.SearchMailboxes($smSearchMailbox);
                Write-Host("Items Remaining : " + $srCol[0].SearchResult.ItemCount);
            } while ($srCol[0].SearchResult.ItemCount-gt 0 );
            
        }            
    }
    Write-Output $rptCollection.Values 
}