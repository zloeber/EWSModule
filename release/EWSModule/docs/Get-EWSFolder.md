---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSFolder
## SYNOPSIS
Return a mailbox folder object.

## SYNTAX

### FolderAsString (Default)
```
Get-EWSFolder [[-EWSService] <ExchangeService>] [[-Mailbox] <String>] [[-FolderPath] <String>]
 [[-FolderObject] <WellKnownFolderName>] [-PublicFolder]
```

### FolderAsObject
```
Get-EWSFolder [[-EWSService] <ExchangeService>] [[-Mailbox] <String>] [[-FolderPath] <String>]
 [[-FolderObject] <WellKnownFolderName>] [-PublicFolder]
```

## DESCRIPTION
Return a mailbox folder object.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSFolder -FolderObject Contacts
```

Description
-----------
Return the Folder object for the currently connected EWSService account of the well known 'contacts' folder
(\[Microsoft.Exchange.WebServices.Data.WellKnownFolderName\]::contacts)

## PARAMETERS

### -EWSService
Exchange web service connection object to use.
The default is using the currently connected session.

```yaml
Type: ExchangeService
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -Mailbox
Mailbox to target.
If none is provided, impersonation is checked and used if possible, otherwise the EWSService object mailbox is targeted.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderPath
Path of folder in the form of /folder1/folder2

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderObject
Well known folder object.

```yaml
Type: WellKnownFolderName
Parameter Sets: (All)
Aliases: 
Accepted values: Calendar, Contacts, DeletedItems, Drafts, Inbox, Journal, Notes, Outbox, SentItems, Tasks, MsgFolderRoot, PublicFoldersRoot, Root, JunkEmail, SearchFolders, VoiceMail, RecoverableItemsRoot, RecoverableItemsDeletions, RecoverableItemsVersions, RecoverableItemsPurges, ArchiveRoot, ArchiveMsgFolderRoot, ArchiveDeletedItems, ArchiveRecoverableItemsRoot, ArchiveRecoverableItemsDeletions, ArchiveRecoverableItemsVersions, ArchiveRecoverableItemsPurges, SyncIssues, Conflicts, LocalFailures, ServerFailures, RecipientCache, QuickContacts, ConversationHistory, ToDoSearch

Required: False
Position: 3
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -PublicFolder
Force target a public folder instead.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

## OUTPUTS

## NOTES
Author: Zachary Loeber
Site: http://www.the-little-things.net/
Requires: Powershell 3.0
Version History
1.0.0 - Initial release

## RELATED LINKS

