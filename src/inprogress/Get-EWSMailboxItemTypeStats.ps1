function Get-EWSMailboxItemTypeStats { 
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
        [Parameter(Position=1, Mandatory=$true)] [String]$ItemType,
        [Parameter(Position=2, Mandatory=$false)] [Switch]$FolderList
    )  
    $service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
    $KQL = ""
    if([string]::IsNullOrEmpty($ItemType))
    {
        $KQL = "kind:email OR kind:meetings OR kind:contacts OR kind:tasks OR kind:notes OR kind:IM OR kind:rssfeeds OR kind:voicemail";
    }
    else{
        $KQL = "kind:" + $ItemType
    }    
    
    if(!$FolderList.IsPresent){
        Get-EWSeDiscoveryKeyWordStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName -Prefix "kind:"
    }
    else{
        New-EWSeDiscoveryPreviewItemsStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName
    }
}