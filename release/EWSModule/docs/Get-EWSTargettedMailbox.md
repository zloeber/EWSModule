---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSTargettedMailbox
## SYNOPSIS
Return the intended targeted mailbox for ews operations.

## SYNTAX

```
Get-EWSTargettedMailbox [[-EWSService] <ExchangeService>] [[-Mailbox] <String>]
```

## DESCRIPTION
Return the intended targeted mailbox for operations.
If an email address string is passed we will try to connect to it with non-impersonation rights.
If the Mailbox parameter is empty or null then we will look at the ews object to see if impersonation is set and return that mailbox if found.
Otherwise
we use the ews object login ID.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSTargetedMailbox -Mailbox jdoe
```

Description
-----------
Reterns the email address jdoe from the domain.

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

## INPUTS

## OUTPUTS

## NOTES
Author: Zachary Loeber
Site: http://www.the-little-things.net/
Requires: Powershell 3.0
Version History
1.0.0 - Initial release

## RELATED LINKS

