---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Import-EWSDll

## SYNOPSIS
Load EWS dlls.

## SYNTAX

```
Import-EWSDll [[-EWSManagedApiPath] <String>]
```

## DESCRIPTION
Load EWS dlls.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Import-EWSDll
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

## INPUTS

## OUTPUTS

## NOTES
Author: Zachary Loeber
Site: http://www.the-little-things.net/
Requires: Powershell 3.0
Version History
1.0.0 - Initial release

## RELATED LINKS

