---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Get-EWSCalenderViewAppointments
## SYNOPSIS
Uses a slower method for accessing and returning calendar appointments

## SYNTAX

```
Get-EWSCalenderViewAppointments [[-EWSService] <ExchangeService>] [[-Mailbox] <String>]
 [[-StartRange] <DateTime>] [[-EndRange] <DateTime>]
```

## DESCRIPTION
Uses a slower method for accessing and returning calendar appointments

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```

```

PS \> 

Description
-----------
TBD

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

### -StartRange
Start of when to look for appointments.

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: (Get-Date)
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndRange
End of when to look for appointments.

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: ((Get-Date).AddMonths(12))
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

