function Get-EWSContact {
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
        [parameter(HelpMessage='Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position=1, HelpMessage='Mailbox to target.')]
        [string]$Mailbox,
        [Parameter(Position=2, Mandatory=$true)]
        [string]$EmailAddress,
        [Parameter(Position=3)]
        [string]$Folder,
        [Parameter(Position=4)]
        [ValidateSet('DirectoryOnly','DirectoryThenContacts','ContactsOnly','ContactsThenDirectory')]
        [ews_resolvenamelocation]$SearchType = 'ContactsThenDirectory',
        [Parameter(Position=5)]
        [switch]$Partial
    )  
    # Pull in all the caller verbose,debug,info,warn and other preferences
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand
    
    if (-not (Get-EWSModuleInitializationState)) {
        throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
    }
    
    if ($EWSService -eq $null) {
        Write-Verbose "$($FunctionName): Using module local ews service object"
        $EWSService = Get-EWSService
    }
    
    if ($EWSService -eq $null) {
        throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
    }

    if( -not [string]::IsNullOrEmpty($Folder) ) {
        $folderid = Get-EWSFolder -EWSService $EWSService -FolderPath $Folder -Mailbox $Mailbox
    }
    else {
        $folderid = Get-EWSFolder -EWSService $EWSService -FolderObject Contacts -Mailbox $Mailbox
    }
    
    # Get the parent folder ID
    $type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
    $type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
    $ParentFolderIds = [Activator]::CreateInstance($type)
    $ParentFolderIds.Add($folderid.Id)

    try {
        $ncCol = @($EWSService.ResolveName($EmailAddress,$ParentFolderIds,$SearchType,$true))
    }
    catch {
        throw "$($FunctionName): Unable to resolve contact $EmailAddress"
    }
    foreach($Result in $ncCol) {
        # If the Contact property is null then this is likely just a user contact
        if($Result.Contact -eq $null) {
            if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -or $Partial){
                # Convert to a full EWS contact and return
                Write-Verbose "$($FunctionName): Returning contact from the mailbox contacts."
                [ews_contact]::Bind($EWSService,$Result.Mailbox.Id)
            }
        }
        # Otherwise it was probably found in the directory.
        else{
            if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -or $Partial){
                if($Result.Mailbox.MailboxType -eq [ews_mailboxtype]::Mailbox){
                    $UserDn = Get-EWSUserDN -EWSService $EWSService -EmailAddress $Result.Mailbox.Address
                    $ncCola = $EWSService.ResolveName($UserDn,$ParentFolderIds,$SearchType,$true)
                    Write-Verbose ("$($FunctionName): Number of matching Contacts Found - " + $ncCola.Count)
                    foreach($aResult in $ncCola){
                        Write-Verbose "$($FunctionName): Returning contact from the directory."
                        $aResult.Contact
                    }
                }
            }
        }
    }
}