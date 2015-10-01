function Get-EWSContactFolder {
    param (
        [Parameter(Position=0, Mandatory=$true)]
        [string]$FolderPath,
        [Parameter(Position=1, Mandatory=$true)]
        [string]$SmptAddress,
        [Parameter(Position=2, Mandatory=$true)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )
    ## Find and Bind to Folder based on Path  
    #Define the path to search should be seperated with \  
    #Bind to the MSGFolder Root  
    $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$SmptAddress)   
    $tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
    #Split the Search path into an array  
    $fldArray = $FolderPath.Split("\") 
     #Loop through the Split Array and do a Search for each level of folder 
    for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
        #Perform search based on the displayname of each folder level 
        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint]) 
        $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
        if ($findFolderResults.TotalCount -gt 0){ 
            foreach($folder in $findFolderResults.Folders){ 
                $tfTargetFolder = $folder                
            } 
        } 
        else{ 
            Write-host ("Error Folder Not Found check path and try again")  
            $tfTargetFolder = $null  
            break  
        }     
    }  
    if($tfTargetFolder -ne $null){
        return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$tfTargetFolder.Id)
    }
    else {
        throw ("Folder Not found")
    }
}