---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Get-EWSRootFolderID

## SYNOPSIS
Return a mailbox folder object.

## SYNTAX

```
Get-EWSRootFolderID [[-EWSService] <ExchangeService>] [[-Mailbox] <String>] [[-FolderBase] <String>]
```

## DESCRIPTION
Return a mailbox folder object.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSRootFolderID -EWSService $EWSService -FolderRoot Contacts -Mailbox 'jdoe@contoso.com'
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

### -FolderBase
A well known folder base name (Inbox, Calendar, Contacts, et cetera..)

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: MsgFolderRoot
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

