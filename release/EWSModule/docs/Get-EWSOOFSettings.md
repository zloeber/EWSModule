---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSOOFSettings
## SYNOPSIS
Get the out of office settings for a mailbox.

## SYNTAX

```
Get-EWSOOFSettings [[-EWSService] <ExchangeService>] [-Mailbox] <String>
```

## DESCRIPTION
Get the out of office settings for a mailbox.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EWSOOFSettings -Mailbox mailbox@domain.com
```

Description
--------------
Get the out of office settings for mailbox@domain.com

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

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: 
Accept pipeline input: True (ByPropertyName)
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

