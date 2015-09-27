function Export-EWSContact {
    <# 
    .SYNOPSIS 
        Exports a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API to a VCF File 
    .DESCRIPTION 
        Exports a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
    .EXAMPLE 
        Example 1 To Export a contact to local file
        Export-EWSContact -MailboxName mailbox@domain.com -EmailAddress address@domain.com -FileName c:\export\filename.vcf
        If the file already exists it will handle creating a unique filename
    .EXAMPLE
        Example 2 To export from a contacts subfolder use
        Export-EWSContact -MailboxName mailbox@domain.com -EmailAddress address@domain.com -FileName c:\export\filename.vcf -folder \contacts\subfolder
    #> 
   [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
        [Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
        [Parameter(Position=3, Mandatory=$true)] [string]$FileName,
        [Parameter(Position=4, Mandatory=$false)] [string]$Folder,
        [Parameter(Position=5, Mandatory=$false)] [switch]$Partial
        
    )
    if($Folder){
        if($Partial.IsPresent){
            $Contacts = Get-EWSContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder -Partial
        }
        else{
            $Contacts = $Contacts = Get-EWSContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials -Folder $Folder
        }
    }
    else{
        if($Partial.IsPresent){
            $Contacts = Get-EWSContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials  -Partial
        }
        else{
            $Contacts = $Contacts = Get-EWSContact -MailboxName $MailboxName -EmailAddress $EmailAddress -Credentials $Credentials 
        }
    }    

    $Contacts | ForEach-Object{
        $Contact = $_
        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)    
          $psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent); 
        $Contact.load($psPropset)
        $FileName = Make-UniqueFileName -FileName $FileName
        [System.IO.File]::WriteAllBytes($FileName,$Contact.MimeContent.Content) 
        write-verbose ("Exported " + $FileName)  
    }
}