---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Set-ServerCertificateValidationCallback
## SYNOPSIS
Sets the current certificate validation callback setting

## SYNTAX

```
Set-ServerCertificateValidationCallback [[-CertCallback] <String>]
```

## DESCRIPTION
Sets the current certificate validation callback setting

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Set-ServerCertificateValidationCallback
```

## PARAMETERS

### -CertCallback
Defaults to \[System.Net.ServicePointManager\]::ServerCertificateValidationCallback

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: [System.Net.ServicePointManager]::ServerCertificateValidationCallback
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

