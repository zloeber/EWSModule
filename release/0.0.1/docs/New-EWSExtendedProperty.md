---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# New-EWSExtendedProperty

## SYNOPSIS
Creates a new extended property which can be assigned to items in outlook.

## SYNTAX

```
New-EWSExtendedProperty [[-PropertyType] <MapiPropertyType[]>] [-PropertyName] <String>
```

## DESCRIPTION
Creates a new extended property which can be assigned to items in outlook.
These are generally hidden to end 
users but can be invaluable in creating items that you can then later locate again and know they were created by
your processes.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$CalEntry = New-EWSCalendarEntry -IsAllDayEvent $true -Subject 'My Event' -Location 'Elsewhere'
```

Description
--------------
Creates a new calendar entry as an all day event called 'My Event' and stores it in $CalEntry

## PARAMETERS

### -PropertyType
Type of extended property to create.

```yaml
Type: MapiPropertyType[]
Parameter Sets: (All)
Aliases: 
Accepted values: ApplicationTime, ApplicationTimeArray, Binary, BinaryArray, Boolean, CLSID, CLSIDArray, Currency, CurrencyArray, Double, DoubleArray, Error, Float, FloatArray, Integer, IntegerArray, Long, LongArray, Null, Object, ObjectArray, Short, ShortArray, SystemTime, SystemTimeArray, String, StringArray

Required: False
Position: 1
Default value: [System.Enum]::GetValues([ews_mapiproptype])
Accept pipeline input: False
Accept wildcard characters: False
```

### -PropertyName
Name of extended property

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: None
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

