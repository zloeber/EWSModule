---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Get-EWSFolderItem

## SYNOPSIS
Returns items from a mailbox folder.

## SYNTAX

```
Get-EWSFolderItem [[-EWSService] <ExchangeService>] [[-Mailbox] <String>] [[-FolderPath] <String>]
 [[-SearchBase] <String>] [[-Count] <Int32>] [[-Filter] <String>]
```

## DESCRIPTION
Returns items from a mailbox folder using either an AQS search filter or just by a number of desired results.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSFolderItem -Count 10 -SearchString 'Subject:Change Notice' -verbose
```

Retrieves all emails containing the term 'Change Notice' in the subject from the Inbox.
Results are fetched 10 at a time and verbose
output is displayed on the screen.

### -------------------------- EXAMPLE 2 --------------------------
```
Get-EWSFolderItem -Count 10 -SearchString 'Received:Yesterday' -verbose
```

Retrieves all emails in the connected user's Inbox received yesterday.
Results are fetched 10 at a time and verbose
output is displayed on the screen.

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
A specific path to search.

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
Base folder to return items from.
The default is Inbox.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: Inbox
Accept pipeline input: False
Accept wildcard characters: False
```

### -Count
Number of items to return if not using a search string.
With a search string it is the number of items returned per page

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: 

Required: False
Position: 5
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filter
Search string in AQS format for returning items.
For complete documentation on syntax see https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 6
Default value: None
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

[https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx](https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx)

