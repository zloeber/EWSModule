---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Set-EWSMailboxImpersonation
## SYNOPSIS
Set the impersonation for a mailbox.

## SYNTAX

```
Set-EWSMailboxImpersonation [[-EWSService] <Object>] [-Mailbox] <String> [-SkipValidation]
```

## DESCRIPTION
Set the impersonation for a mailbox.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Set-EWSMailboxImpersonation -Mailbox mailbox@domain.com
```

Description
--------------
Set impersonation mode for the current connected EWS user for mailbox@domain.com

## PARAMETERS

### -EWSService
Exchange web service connection object to use.
The default is using the currently connected session.

```yaml
Type: Object
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

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipValidation
Do not validate if you have impersonation rights for the mailbox (can speed things up quite a bit)

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
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

