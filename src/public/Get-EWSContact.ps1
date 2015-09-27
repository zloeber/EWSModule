﻿function Get-EWSContact {
    <# 
    .SYNOPSIS 
        Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
    .DESCRIPTION 
        Gets a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
    .EXAMPLE
        Example 1 To get a Contact from a Mailbox's default contacts folder
        Get-EWSContact -MailboxName mailbox@domain.com -EmailAddress contact@email.com
    .EXAMPLE    
        Example 2  The Partial Switch can be used to do partial match searches. Eg to return all the contacts that contain a particular word (note this could be across all the properties that are searched) you can use
        Get-EWSContact -MailboxName mailbox@domain.com -EmailAddress glen -Partial
    .EXAMPLE
        Example 3 By default only the Primary Email of a contact is checked when you using ResolveName if you want it to search the multivalued Proxyaddressses property you need to use something like the following
        Get-EWSContact -MailboxName  mailbox@domain.com -EmailAddress smtp:info@domain.com -Partial
    .EXAMPLE
        Example 4 Or to search via the SIP address you can use
        Get-EWSContact -MailboxName  mailbox@domain.com -EmailAddress sip:info@domain.com -Partial
    #>
   [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
        [Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=3, Mandatory=$false)] [string]$Folder,
        [Parameter(Position=4, Mandatory=$false)] [switch]$SearchGal,
        [Parameter(Position=5, Mandatory=$false)] [switch]$Partial,
        [Parameter(Position=6, Mandatory=$false)] [switch]$useImpersonation
    )  
    $service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
    if($useImpersonation.IsPresent){
        $service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
    }
    $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
    if($SearchGal)
    {
        $Error.Clear();
        $ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly,$true);
        if($Error.Count -eq 0){
            foreach($Result in $ncCol){    
                if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
                    Write-Output $ncCol.Contact
                }
                else{
                    Write-host -ForegroundColor Yellow ("Partial Match found but not returned because Primary Email Address doesn't match consider using -Partial " + $ncCol.Contact.DisplayName + " : Subject-" + $ncCol.Contact.Subject + " : Email-" + $Result.Mailbox.Address)
                }
            }
        }
    }
    else
    {
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
                    $ResultWritten = $false
                    foreach($Result in $ncCol){
                        if($Result.Contact -eq $null){
                            if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
                                $Contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$Result.Mailbox.Id)
                                Write-Output $Contact  
                                $ResultWritten = $true
                            }
                        }
                        else{
                        
                            if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
                                if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
                                    $ResultWritten = $true
                                    $UserDn = Get-EWSUserDN -EmailAddress $Result.Mailbox.Address -Credentials $Credentials 
                                    $ncCola = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true);
                                    if ($ncCola.Count -eq 0) {  
                                        #Write-Host -ForegroundColor Yellow ("No Contact Found")            
                                    }
                                    else
                                    {
                                        $ResultWritten = $true
                                        Write-Host ("Number of matching Contacts Found " + $ncCola.Count)
                                        foreach($aResult in $ncCola){
                                            $Contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$aResult.Mailbox.Id)
                                            Write-Output $Contact
                                        }
                                        
                                    }
                                }
                            }
                        }
                        
                    }
                    if(!$ResultWritten){
                        Write-Host -ForegroundColor Yellow ("No Contract Found")
                    }
                }
            }

            
        }
    }
}