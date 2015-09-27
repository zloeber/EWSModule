function Get-EWSMailboxItemStats {
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
        [Parameter(Position=2, Mandatory=$false)] [string]$KQL,
        [Parameter(Position=3, Mandatory=$false)] [DateTime]$Start,
        [Parameter(Position=4, Mandatory=$false)] [DateTime]$End,
        [Parameter(Position=5, Mandatory=$false)] [Switch]$FolderList
        
    )  
    $service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials            
    if((![string]::IsNullOrEmpty($Start)) -band (![string]::IsNullOrEmpty($End))){
        $KQL = "Received:" + $Start.ToString("yyyy-MM-dd") + ".." + $End.ToString("yyyy-MM-dd")
    }
    if(!$FolderList.IsPresent){
        Get-EWSeDiscoveryKeyWordStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName -Prefix "kind:"
    }
    else{
        New-EWSeDiscoveryPreviewItemsStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName
    }
}