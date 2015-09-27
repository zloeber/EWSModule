function Copy-EWSContactsFromGalToMailbox {
    <# 
    .SYNOPSIS 
     Copies a Contact from the Global Address List to a Local Mailbox Contacts folder using the  Exchange Web Services API  
     
    .DESCRIPTION 
      Copies a Contact from the Global Address List to a Local Mailbox Contacts folder using the  Exchange Web Services API
      
      Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
    .EXAMPLE 
        To Copy a Gal contacts to local Contacts folder
        Copy-EWSContactsFromGalToMailbox -MailboxName mailbox@domain.com -EmailAddress email@domain.com  
    .EXAMPLE
        Copy a GAL contact to a Contacts subfolder
        Copy-EWSContactsFromGalToMailbox -MailboxName mailbox@domain.com -EmailAddress email@domain.com  -Folder \Contacts\UnderContacts
    #> 
   [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
        [Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=3, Mandatory=$false)] [string]$Folder,
        [Parameter(Position=4, Mandatory=$false)] [switch]$IncludePhoto,
        [Parameter(Position=5, Mandatory=$false)] [switch]$useImpersonation
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
    $Error.Clear();
    $ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly,$true);
    if($Error.Count -eq 0){
        foreach($Result in $ncCol){                
            if($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()){                    
                $type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
                $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
                $ParentFolderIds = [Activator]::CreateInstance($type)
                $ParentFolderIds.Add($Contacts.Id)
                $Error.Clear();
                $ncCola = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true);
                $createContactOkay = $false
                if($Error.Count -eq 0){
                    if ($ncCola.Count -eq 0) {                            
                        $createContactOkay = $true;    
                    }
                    else{
                        foreach($aResult in $ncCola){
                            if($aResult.Contact -eq $null){
                                Write-host "Contact already exists " + $aResult.Contact.DisplayName
                                throw ("Contact already exists")
                            }
                            else{
                                if((Validate-EmailAddres -EmailAddress $Result.Mailbox.Address)){
                                    if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
                                        $UserDn = Get-EWSUserDN -Credentials $Credentials -EmailAddress $Result.Mailbox.Address
                                        $ncColb = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true);
                                        if ($ncColb.Count -eq 0) {  
                                            $createContactOkay = $true;        
                                        }
                                        else
                                        {
                                            Write-Host -ForegroundColor  Red ("Number of existing Contacts Found " + $ncColb.Count)
                                            foreach($Result in $ncColb){
                                                Write-Host -ForegroundColor  Red ($ncColb.Mailbox.Name)
                                            }
                                            throw ("Contact already exists")
                                        }
                                    }
                                }
                                else{
                                    Write-Host -ForegroundColor Yellow ("Email Address is not valid for GAL match")
                                }
                            }
                        }
                    }
                    if($createContactOkay){
                        #check for SipAddress
                        $IMAddress = ""
                        if($ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] -ne $null){
                            $email1 = $ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address
                            if($email1.tolower().contains("sip:")){
                                $IMAddress = $email1
                            }
                        }
                        if($ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] -ne $null){
                            $email2 = $ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Address
                            if($email2.tolower().contains("sip:")){
                                $IMAddress = $email2
                                $ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = $null
                            }
                        }
                        if($ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3] -ne $null){
                            $email3 = $ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address
                            if($email3.tolower().contains("sip:")){
                                $IMAddress = $email3
                                $ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3] = $null
                            }
                        }
                        if($IMAddress -ne ""){
                            $ncCol.Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress
                        }    
                        $ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = $null
                        $ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3] = $null
                        $ncCol.Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address = $ncCol.Mailbox.Address.ToLower()
                        $ncCol.Contact.FileAs = $ncCol.Contact.DisplayName
                        if($IncludePhoto){                    
                            $PhotoURL = AutoDiscoverPhotoURL -EmailAddress $MailboxName  -Credentials $Credentials
                            $PhotoSize = "HR120x120" 
                            $PhotoURL= $PhotoURL + "/GetUserPhoto?email="  + $ncCol.Mailbox.Address + "&size=" + $PhotoSize;
                            $wbClient = new-object System.Net.WebClient
                            $creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString()) 
                            $wbClient.Credentials = $creds
                            $photoBytes = $wbClient.DownloadData($PhotoURL);
                            $fileAttach = $ncCol.Contact.Attachments.AddFileAttachment("contactphoto.jpg",$photoBytes)
                            $fileAttach.IsContactPhoto = $true
                        }
                        $ncCol.Contact.Save($Contacts.Id);
                        Write-Host ("Contact copied")
                    }
                }
            }
        }
    }
}