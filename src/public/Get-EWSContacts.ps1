function Get-EWSContacts {
    <# 
    .SYNOPSIS 
        Gets all contacts in a Contact folder in a Mailbox using the Exchange Web Services API 
     
    .DESCRIPTION
        Gets all contacts in a Contact folder in a Mailbox using the Exchange Web Services API 
      
    .EXAMPLE
        To get all contacts from a Mailbox's default contacts folder
        Get-EWSContacts -Mailbox mailbox@domain.com
    .EXAMPLE
        To get all the Contacts from subfolder of the Mailbox's default contacts folder
        Get-EWSContacts -Mailbox mailbox@domain.com -Folder \Contact\test
    #>
    [CmdletBinding()] 
    param(
        [parameter(HelpMessage = 'Connected EWS object.')]
        [ews_service]$EWSService,
        [parameter(Position = 1, HelpMessage = 'Mailbox to target.')]
        [string]$Mailbox,
        [Parameter(Position = 2, Mandatory = $true)]
        [string]$EmailAddress,
        [Parameter(Position = 3)]
        [string]$Folder,
        [Parameter(Position = 4)]
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

    $SfSearchFilter = New-Object ews_searchfilter_isequalto([ews_schema_item]::ItemClass,"IPM.Contact")
    $ivItemView =  New-Object ews_itemview(1000)
    $fiItems = $null

    do {
        $fiItems = $EWSService.FindItems($folderid.Id,$SfSearchFilter,$ivItemView)
        if($fiItems.Items.Count -gt 0) {
            $psPropset = new-object ews_propset([ews_basepropset]::FirstClassProperties)
            [Void]$EWSService.LoadPropertiesForItems($fiItems,$psPropset)
            foreach($Item in $fiItems.Items) {
                Write-Output $Item
            }
        }
        $ivItemView.Offset += $fiItems.Items.Count    
    } while($fiItems.MoreAvailable -eq $true) 
}