---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Install-EWSDll

## SYNOPSIS
Attempts to download and extract the ews dll needed for this library.

## SYNTAX

```
Install-EWSDll [[-source] <String>] [[-destination] <String>] [-SkipDownload]
```

## DESCRIPTION
Attempts to download and extract the ews dll needed for this library.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Install-EWSDll
```

Description
--------------
Attempts to download and extract the appropriate DLL for EWS from http://www.microsoft.com/en-us/download/details.aspx?id=28952

## PARAMETERS

### -source
Web URL  to the EWSmanagedApi.msi file

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: Http://download.microsoft.com/download/8/9/9/899EEF2C-55ED-4C66-9613-EE808FCF861C/EwsManagedApi.msi
Accept pipeline input: False
Accept wildcard characters: False
```

### -destination
Destination for the extracted DLL

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipDownload
Extract the file from the EWSmanagedApi.msi file you have already pre-downloaded to .\EWSFiles\

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

## INPUTS

## OUTPUTS

## NOTES
Author: Zachary Loeber
Site: http://www.the-little-things.net/
Requires: Powershell 3.0
Version History
1.0.0 - Initial release

## RELATED LINKS

