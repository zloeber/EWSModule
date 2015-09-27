function Get-EWSAttachmentTypeMailboxStats { 
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
        [Parameter(Position=2, Mandatory=$false)] [PSObject]$AttachmentList,
        [Parameter(Position=3, Mandatory=$false)] [string]$AttachmentType,
        [Parameter(Position=4, Mandatory=$false)] [string]$AttachmentName,
        [Parameter(Position=5, Mandatory=$false )] [Switch]$FolderList
    )  
    $service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
    $KQL = ""
    if([string]::IsNullOrEmpty($AttachmentList))
    {            
        $AttachmentList = @()
        $AttachmentList += "xlsx"
        $AttachmentList += "docx"
        $AttachmentList += "doc"
        $AttachmentList += "xls"
        $AttachmentList += "pptx"
        $AttachmentList += "ppt"
        $AttachmentList += "txt"
        $AttachmentList += "mp3"
        $AttachmentList += "zip"
        $AttachmentList += "txt"
        $AttachmentList += "wma"
        $AttachmentList += "pdf"
    }
    if(![string]::IsNullOrEmpty($AttachmentType))
    {
        $AttachmentList = @()
        $AttachmentList += $AttachmentType
    }
    else
    {
        if(![string]::IsNullOrEmpty($AttachmentName))
        {
            $KQL = "Attachment:" + $AttachmentName
        }
    }            
    if($KQL -eq ""){
        foreach($Item in $AttachmentList){
            if($KQL -eq ""){
                $KQL = "Attachment:." + $Item
            }
            else{
                $KQL += " OR Attachment:." + $Item
            }
        }            
    }
    if(!$FolderList.IsPresent)
    {
        Get-EWSeDiscoveryKeyWordStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName -Prefix "Attachment:."
    }
    else
    {            
        New-EWSeDiscoveryPreviewItemsStats -service $service -KQL $KQL -SearchableMailboxString $MailboxName
    }
}
