---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Get-EmailAddressFromAD

## SYNOPSIS
Return the email address of a User ID from AD.

## SYNTAX

```
Get-EmailAddressFromAD [[-UserID] <String>]
```

## DESCRIPTION
Return the email address of a User ID from AD.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Get-EmailAddressFromAD -UserID jdoe
```

Reterns the email address for jdoe from the domain.

## PARAMETERS

### -UserID
User ID to search for in AD.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
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

