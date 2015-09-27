function Get-EWSContacts {
    <# 
    .SYNOPSIS 
        Gets contacts in a Contact folder in a Mailbox using the  Exchange Web Services API 
     
    .DESCRIPTION 
        Gets contacts in a Contact folder in a Mailbox using the  Exchange Web Services API 
      
    .EXAMPLE
        To get a Contact from a Mailbox's default contacts folder
        Get-EWSContacts -MailboxName mailbox@domain.com
    .EXAMPLE
        To get all the Contacts from subfolder of the Mailbox's default contacts folder
        Get-EWSContacts -MailboxName mailbox@domain.com -Folder \Contact\test
        
    #>
   [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=2, Mandatory=$false)] [string]$Folder,
        [Parameter(Position=3, Mandatory=$false)] [switch]$useImpersonation
    )
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
        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Contact") 
        #Define ItemView to retrive just 1000 Items    
        $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
        $fiItems = $null    
        do {    
            $fiItems = $service.FindItems($Contacts.Id,$SfSearchFilter,$ivItemView)    
            if($fiItems.Items.Count -gt 0){
                $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
                [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
                foreach($Item in $fiItems.Items){      
                    Write-Output $Item    
                }
            }
            $ivItemView.Offset += $fiItems.Items.Count    
        } while($fiItems.MoreAvailable -eq $true) 
    }
}