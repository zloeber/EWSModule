---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Get-EWSContacts

## SYNOPSIS
Gets all contacts in a Contact folder in a Mailbox using the Exchange Web Services API

## SYNTAX

```
Get-EWSContacts [-EWSService <ExchangeService>] [[-Mailbox] <String>] [-EmailAddress] <String>
 [[-Folder] <String>] [[-SearchType] <ResolveNameSearchLocation>] [-Partial]
```

## DESCRIPTION
Gets all contacts in a Contact folder in a Mailbox using the Exchange Web Services API

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSContacts -Mailbox mailbox@domain.com
```

Get all contacts from a Mailbox's default contacts folder

### -------------------------- EXAMPLE 2 --------------------------
```
Get-EWSContacts -Mailbox mailbox@domain.com -Folder \Contact\test
```

Get all the Contacts from subfolder of the Mailbox's default contacts folder

## PARAMETERS

### -EWSService
Exchange web service connection object to use.
The default is using the currently connected session.

```yaml
Type: ExchangeService
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
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

### -EmailAddress
Email address of the contact to search.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 3
Default value: None
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
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SearchType
Search type determines different orders to search.
The default is ContactsThenDirectory

```yaml
Type: ResolveNameSearchLocation
Parameter Sets: (All)
Aliases: 
Accepted values: DirectoryOnly, DirectoryThenContacts, ContactsOnly, ContactsThenDirectory

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
Requires: Powershell 3.0
Version History
1.0.0 - Initial release

## RELATED LINKS

[http://www.the-little-things.net/](http://www.the-little-things.net/)

[https://www.github.com/zloeber/EWSModule](https://www.github.com/zloeber/EWSModule)

