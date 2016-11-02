---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Remove-EWSCalendarAppointment

## SYNOPSIS
Remove a calendar appointment object from a mailbox.

## SYNTAX

```
Remove-EWSCalendarAppointment [[-EWSService] <Object>] [-Appointment] <Appointment> [[-DeleteMode] <String>]
 [[-CancellationMode] <String>]
```

## DESCRIPTION
Remove a calendar appointment object from a mailbox.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Remove-EWSCalendarAppointment -Appointment $Appt -DeleteMode 'SoftDelete' -CancellationMode 'SendToAllAndSaveCopy'
```

Removes the calendar appointment stored in $Appt using softdelete and sending a cancellation notice and saving a copy of the sent cancellation.

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
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Appointment
EWS Calendar appointment object

```yaml
Type: Appointment
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -DeleteMode
Method of deletion for the appointment.
Can be 'HardDelete','SoftDelete', or 'MoveToDeletedItems'

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: HardDelete
Accept pipeline input: False
Accept wildcard characters: False
```

### -CancellationMode
How cancellation notices will be sent upon deletion.
Can be 'SendToNone','SendOnlyToAll','SendOnlyToChanged','SendToAllAndSaveCopy', or 'SendToChangedAndSaveCopy'

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: SendToNone
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

