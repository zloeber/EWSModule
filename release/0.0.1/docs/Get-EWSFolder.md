---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Get-EWSFolder

## SYNOPSIS
Return a mailbox folder object.

## SYNTAX

```
Get-EWSFolder [[-EWSService] <ExchangeService>] [[-Mailbox] <String>] [[-FolderPath] <String>]
 [[-SearchBase] <String>]
```

## DESCRIPTION
Return a mailbox folder object.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSFolder -FolderRoot Contacts
```

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
Default value: None
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
Default value: None
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
Default value: \
Accept pipeline input: False
Accept wildcard characters: False
```

### -SearchBase
Well known folder object.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: MsgFolderRoot
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

## OUTPUTS

## NOTES
Author: Zachary Loeber
Requires: Powershell 3.0
Version History
1.0.0 - Initial release

## RELATED LINKS

[http://www.the-little-things.net/](http://www.the-little-things.net/)

[https://www.github.com/zloeber/EWSModule](https://www.github.com/zloeber/EWSModule)

