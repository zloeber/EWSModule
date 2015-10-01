function Delete-EWSContact {
    <# 
    .SYNOPSIS 
     Deletes a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
     
    .DESCRIPTION 
      Deletes a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
      
      Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
    .EXAMPLE
        Example 1 To delete a contact from the default contacts folder
        Delete-EWSContact -MailboxName mailbox@domain.com -EmailAddress email@domain.com 
    .EXAMPLE
        Example2 To delete a contact from a non user subfolder
        Delete-EWSContact -MailboxName mailbox@domain.com -EmailAddress email@domain.com -Folder \Contacts\Subfolder
    #> 
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
        [Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=3, Mandatory=$false)] [switch]$force,
        [Parameter(Position=4, Mandatory=$false)] [string]$Folder,
        [Parameter(Position=5, Mandatory=$false)] [switch]$Partial
    )  
    $service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
    if($useImpersonation.IsPresent){
        $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
    }
    $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
    if($Folder){
        $Contacts = Get-EWSContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
    }
    else{
        $Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
    }
    if($service.URL){
        $type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
        $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
        $ParentFolderIds = [Activator]::CreateInstance($type)
        $ParentFolderIds.Add($Contacts.Id)
        $Error.Clear();
        $ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true);
        if($Error.Count -eq 0){
            if ($ncCol.Count -eq 0) {
                Write-Host -ForegroundColor Yellow ("No Contact Found")        
            }
            else{
                foreach($Result in $ncCol){
                    if($Result.Contact -eq $null){
                        $contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$Result.Mailbox.Id) 
                        if($force){
                            if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower())){
                                $contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)  
                                Write-Host ("Contact Deleted " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address)
                            }
                            else
                            {
                                Write-Host ("This script won't allow you to force the delete of partial matches")
                            }
                        }
                        else{
                            if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
                                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""  
                                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","" 
                               
                                $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)  
                                $message = "Do you want to Delete contact with DisplayName " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address
                                $result = $Host.UI.PromptForChoice($caption,$message,$choices,1)  
                                if($result -eq 0) {                       
                                    $contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) 
                                    Write-Host ("Contact Deleted")
                                } 
                                else{
                                    Write-Host ("No Action Taken")
                                }
                            }
                            
                        }
                    }
                    else{
                        if((Validate-EmailAddres -EmailAddress $Result.Mailbox.Address)){
                            if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
                                $UserDn = Get-EWSUserDN -Credentials $Credentials -EmailAddress $Result.Mailbox.Address
                                $ncCola = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true);
                                if ($ncCola.Count -eq 0) {  
                                    Write-Host -ForegroundColor Yellow ("No Contact Found")            
                                }
                                else
                                {
                                    Write-Host ("Number of matching Contacts Found " + $ncCola.Count)
                                    $rtCol = @()
                                    foreach($aResult in $ncCola){
                                        if(($aResult.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
                                            $contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$aResult.Mailbox.Id) 
                                            if($force){
                                                if($aResult.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()){
                                                    $contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)  
                                                    Write-Host ("Contact Deleted " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address)
                                                }
                                                else
                                                {
                                                    Write-Host ("This script won't allow you to force the delete of partial matches")
                                                }
                                            }
                                            else{
                                                $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""  
                                                $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","" 
                                                $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)  
                                                $message = "Do you want to Delete contact with DisplayName " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address 
                                                $result = $Host.UI.PromptForChoice($caption,$message,$choices,1)  
                                                if($result -eq 0) {                       
                                                    $contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) 
                                                    Write-Host ("Contact Deleted ")
                                                } 
                                                else{
                                                    Write-Host ("No Action Taken")
                                                }
                                                
                                            }
                                        }
                                        else{
                                            Write-Host ("Skipping Matching because email address doesn't match address on match " + $aResult.Mailbox.Address.ToLower())
                                        }
                                    }                                
                                }
                            }
                        }
                        else
                        {
                            Write-Warning -ForegroundColor Yellow ("Email Address is not valid for GAL match")
                        }
                    }
                }
            }
        }    
        
    }
}