function Get-EWSMailboxConversationStats { 
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,        
        [Parameter(Position=3, Mandatory=$true)] [PSObject]$ParticipantList,
        [Parameter(Position=4, Mandatory=$false)] [Switch]$FolderList
    )  
    $service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
    $KQL = ""
    foreach($Item in $ParticipantList){
        if($KQL -eq ""){
            $KQL = "Participants:" + $Item
        }
        else{
            $KQL += " OR Participants:" + $Item
        }
    }
    if(!$FolderList.IsPresent){
        Get-EWSeDiscoveryKeyWordStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName -Prefix "kind:"
    }
    else{
        New-EWSeDiscoveryPreviewItemsStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName
    }
}