function Get-EWSContact {
    <# 
    .SYNOPSIS 
    Gets a single contact in a Contact folder within a mailbox using the Exchange Web Services API 
    .DESCRIPTION 
    Gets a single contact in a Contact folder within a mailbox using the Exchange Web Services API
    .PARAMETER EWSService
    Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Mailbox
    Mailbox to target. If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.
    .PARAMETER EmailAddress
    Email address of the contact to search.
    .PARAMETER Folder
    Folder in the mailbox in which the contact is to be searched
    .PARAMETER SearchType
    Search type determines different orders to search. The default is ContactsThenDirectory
    .PARAMETER Partial
    Non-exact match searching.
    .EXAMPLE
    PS> Get-EWSContact -Mailbox mailbox@domain.com -EmailAddress contact@email.com

    Get a Contact from a Mailbox's default contacts folder
    .LINK
    http://www.the-little-things.net/

    .LINK
    https://www.github.com/zloeber/EWSModule

    .NOTES
    Author: Zachary Loeber
    Requires: Powershell 3.0
    Version History
    1.0.0 - Initial release
    #>
    [CmdletBinding()] 
    param(
        [parameter(Position = 0)]
        [ews_service]$EWSService,
        [parameter(Position = 1)]
        [string]$Mailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [string]$EmailAddress,
        [Parameter(Position = 3)]
        [string]$Folder,
        [Parameter(Position = 4)]
        [ValidateSet('DirectoryOnly','DirectoryThenContacts','ContactsOnly','ContactsThenDirectory')]
        [string]$SearchType = 'ContactsThenDirectory',
        [Parameter(Position=5)]
        [switch]$Partial
    )  
    # Pull in all the caller verbose,debug,info,warn and other preferences
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    $FunctionName = $MyInvocation.MyCommand.Name
    
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