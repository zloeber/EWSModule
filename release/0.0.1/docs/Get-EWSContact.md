---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSContact
## SYNOPSIS
Gets a single contact in a Contact folder within a mailbox using the Exchange Web Services API

## SYNTAX

```
Get-EWSContact [[-EWSService] <ExchangeService>] [[-Mailbox] <String>] [-EmailAddress] <String>
 [[-Folder] <String>] [[-SearchType] <String>] [-Partial]
```

## DESCRIPTION
Gets a single contact in a Contact folder within a mailbox using the Exchange Web Services API

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSContact -Mailbox mailbox@domain.com -EmailAddress contact@email.com
```

Description
--------------
Get a Contact from a Mailbox's default contacts folder

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

### -EmailAddress
Email address of the contact to search.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 3
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -Folder
Folder in the mailbox in which the contact is to be searched

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

### -SearchType
Search type determines different orders to search.
The default is ContactsThenDirectory

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 5
Default value: ContactsThenDirectory
Accept pipeline input: False
Accept wildcard characters: False
```

### -Partial
Non-exact match searching.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: 6
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

