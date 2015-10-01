function Create-EWSContactGroup {
    <# 
    .SYNOPSIS 
     Creates a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API 
     
    .DESCRIPTION 
      Creates a Contact Group in a Contact folder in a Mailbox using the  Exchange Web Services API 
      
      Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
    .EXAMPLE
        To create a Contact Group in the default contacts folder 
        Create-EWSContactGroup  -Mailboxname mailbox@domain.com -GroupName GroupName -Members ("member1@domain.com","member2@domain.com")
    .EXAMPLE
        To create a Contact Group in a subfolder of default contacts folder 
        Create-EWSContactGroup  -Mailboxname mailbox@domain.com -GroupName GroupName -Folder \Contacts\Folder1 -Members ("member1@domain.com","member2@domain.com")
    #>
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=2, Mandatory=$false)] [string]$Folder,
        [Parameter(Position=3, Mandatory=$true)] [string]$GroupName,
        [Parameter(Position=4, Mandatory=$true)] [PsObject]$Members,
        [Parameter(Position=5, Mandatory=$false)] [switch]$useImpersonation
    )  
    
    #Connect
    $service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
    if($useImpersonation.IsPresent){
        $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
    }
    $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
    if($Folder){
        $Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
    }
    else{
        $Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
    }
    if($service.URL){
        $ContactGroup = New-Object Microsoft.Exchange.WebServices.Data.ContactGroup -ArgumentList $service
        $ContactGroup.DisplayName = $GroupName
        foreach($Member in $Members){
            $ContactGroup.Members.Add($Member)
        }
        $ContactGroup.Save($Contacts.Id)
        Write-Verbose ("Contact Group created " + $GroupName)
    }
}