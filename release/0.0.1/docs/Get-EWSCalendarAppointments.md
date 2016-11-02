---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Get-EWSCalendarAppointments

## SYNOPSIS
Uses the much faster FindItems as opposed to FindAppointments to return calendar appointments.

## SYNTAX

```
Get-EWSCalendarAppointments [[-EWSService] <ExchangeService>] [-Mailbox <String>] [-FolderPath <String>]
 [-Subject <String>] [-StartsAfter <DateTime>] [-StartsBefore <DateTime>] [-EndsAfter <DateTime>]
 [-EndsBefore <DateTime>] [-CreatedBefore <DateTime>] [-CreatedAfter <DateTime>]
 [-LastOccurrenceBefore <DateTime>] [-LastOccurrenceAfter <DateTime>] [-IsRecurring]
 [-ExtendedProperties <ExtendedPropertyDefinition[]>]
```

## DESCRIPTION
Uses the much faster FindItems as opposed to FindAppointments to return calendar appointments.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```

```

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
Position: Named
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
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
Subject of the appointment(s) being searched

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartsAfter
Start date for the appointment(s) must be after this date

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartsBefore
Start date for the appointment(s) must be before this date

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndsAfter
nd date for the appointment(s) must be after this date

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndsBefore
nd date for the appointment(s) must be before this date

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CreatedBefore
Only appointments created before the given date will be returned

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CreatedAfter
Only appointments created after the given date will be returned

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LastOccurrenceBefore
Only recurring appointments with a last occurrence date before the given date will be returned

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LastOccurrenceAfter
Only recurring appointments with a last occurrence date after the given date will be returned

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IsRecurring
If this switch is present, only recurring appointments are returned

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExtendedProperties
Filter results by custom extended properties.

```yaml
Type: ExtendedPropertyDefinition[]
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
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

