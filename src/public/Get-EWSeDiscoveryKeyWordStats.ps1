function Get-EWSeDiscoveryKeyWordStats {
    [CmdletBinding()]
    param(
        [Parameter(Position=0, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [Parameter(Position=1, Mandatory=$true)] [String]$KQL,
        [Parameter(Position=2, Mandatory=$true)] [String]$SearchableMailboxString,
        [Parameter(Position=3, Mandatory=$true)] [String]$Prefix
    )
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
    $smSearchMailbox.PageSize = 100;
    $smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
    $smSearchMailbox.PerformDeduplication = $false;           
    $smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::StatisticsOnly;
    $srCol = $service.SearchMailboxes($smSearchMailbox);
    $rptCollection = @()
    if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
    {
        foreach($KeyWorkdStat in $srCol[0].SearchResult.KeywordStats){
            if($KeyWorkdStat.Keyword.Contains(" OR ") -eq $false){
                $rptObj = "" | Select Name,ItemHits,Size
                $rptObj.Name = $KeyWorkdStat.Keyword.Replace($Prefix,"")
                $rptObj.Name = $rptObj.Name.Replace($Prefix.ToLower(),"")
                $rptObj.ItemHits = $KeyWorkdStat.ItemHits
                $rptObj.Size = [System.Math]::Round($KeyWorkdStat.Size /1024/1024,2)
                $rptCollection += $rptObj
            }
        }   
    }
    Write-Output $rptCollection
}