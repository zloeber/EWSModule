function ConvertTo-HexId {    
	param (
	        $EWSid,
            $EmailAddress
		  )
	process{
	    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
	    $aiItem.Mailbox = $EmailAddress
	    $aiItem.UniqueId = $EWSid   
	    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EWSId   
	    $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId) 
		return $convertedId.UniqueId
	}
}