---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Initialize-EWS

## SYNOPSIS
Load EWS dlls and create type accelerators for other functions.

## SYNTAX

```
Initialize-EWS [[-EWSManagedApiPath] <String>] [-Uninitialize]
```

## DESCRIPTION
Load EWS dlls and create type accelerators for other functions.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Initialize-EWS
```

## PARAMETERS

### -EWSManagedApiPath
Full path to Microsoft.Exchange.WebServices.dll.
If not provided we will try to load it from several best guess locations.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Uninitialize
Remove previously added type-accelerators.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
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

