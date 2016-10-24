---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSFolderPaths
## SYNOPSIS
Return a mailbox folder object.

## SYNTAX

```
Get-EWSFolderPaths [[-EWSService] <ExchangeService>] [-RootFolderId] <FolderId> [-FolderCache] <PSObject>
 [[-FolderPrefix] <String>]
```

## DESCRIPTION
Return a mailbox folder object.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSFolderPaths
```

Description
-----------
Gets the paths of the currently connected mailbox.

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

### -RootFolderId
Folder to target.
Can target specific mailboxes with Get-EWSFolder

```yaml
Type: FolderId
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderCache
Mailbox foldercache object

```yaml
Type: PSObject
Parameter Sets: (All)
Aliases: 

Required: True
Position: 3
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -FolderPrefix
I forget what this one does, you almost never have to pass it though.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: 
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

