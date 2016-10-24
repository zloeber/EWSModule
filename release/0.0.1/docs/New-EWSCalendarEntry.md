---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# New-EWSCalendarEntry
## SYNOPSIS
Creates an appointment object that can be manipulated or saved.

## SYNTAX

```
New-EWSCalendarEntry [[-EWSService] <ExchangeService>] [[-FreeBusyStatus] <LegacyFreeBusyStatus>]
 [[-IsAllDayEvent] <Boolean>] [[-IsReminderSet] <Boolean>] [[-Start] <DateTime>] [[-End] <DateTime>]
 [[-Subject] <String>] [[-Location] <String>] [[-Body] <String>]
```

## DESCRIPTION
Creates an appointment object that can be manipulated or saved.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$CalEntry = New-EWSCalendarEntry -IsAllDayEvent $true -Subject 'My Event' -Location 'Elsewhere'
```

Description
--------------
Creates a new calendar entry as an all day event called 'My Event' and stores it in $CalEntry

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

### -FreeBusyStatus
FreeBusy status for the appointment.
Can be 'Free','Tentative','Busy','OOF','WorkingElsewhere', or 'NoData'.
Defaults to 'Free'.

```yaml
Type: LegacyFreeBusyStatus
Parameter Sets: (All)
Aliases: 
Accepted values: Free, Tentative, Busy, OOF, WorkingElsewhere, NoData

Required: False
Position: 2
Default value: Free
Accept pipeline input: False
Accept wildcard characters: False
```

### -IsAllDayEvent
Set the flag to mark the appointment as an all day event.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -IsReminderSet
Set the flag to mark the appointment to have a default reminder.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Start
Start time of the appointment.

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: 5
Default value: (Get-Date)
Accept pipeline input: False
Accept wildcard characters: False
```

### -End
End time of the appointment.

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases: 

Required: False
Position: 6
Default value: (Get-Date)
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subject
Appointment subject line.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 7
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -Location
Appointment location.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 8
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -Body
Body of the appointment.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 9
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

