---
external help file: EWSModule-help.xml
online version: 
schema: 2.0.0
---

# Connect-EWS

## SYNOPSIS
Connects to Exchange Web Services.

## SYNTAX

### Default (Default)
```
Connect-EWS [-ExchangeVersion <String>] [-EwsUrl <String>] [-EWSTracing] [-IgnoreSSLCertificate]
```

### CredentialString
```
Connect-EWS -UserName <String> -Password <String> [-Domain <String>] [-ExchangeVersion <String>]
 [-EwsUrl <String>] [-EWSTracing] [-IgnoreSSLCertificate]
```

### CredentialObject
```
Connect-EWS -Credential <PSCredential> [-ExchangeVersion <String>] [-EwsUrl <String>] [-EWSTracing]
 [-IgnoreSSLCertificate]
```

## DESCRIPTION
Connects to Exchange Web Services.
Allows for multiple methods to connect, including autodiscover.
Note that your login ID must be in email format for autodiscover to function.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
$credentials = Get-Credential
```

PS \> Connect-EWS -Creds $credentials -ExchangeVersion 'Exchange2013_SP1' -EwsUrl 'https://webmail.contoso.com/ews/Exchange.asmx'

Description
-----------
Connects to Exchange web services with credentials provided at the prompt.

## PARAMETERS

### -UserName
Username to connect with.

```yaml
Type: String
Parameter Sets: CredentialString
Aliases: 

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password to connect with.

```yaml
Type: String
Parameter Sets: CredentialString
Aliases: 

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Domain
Domain to connect to.

```yaml
Type: String
Parameter Sets: CredentialString
Aliases: 

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Credential
Credential object to connect with.

```yaml
Type: PSCredential
Parameter Sets: CredentialObject
Aliases: Creds

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExchangeVersion
Version of Exchange to target.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: Exchange2010_SP2
Accept pipeline input: False
Accept wildcard characters: False
```

### -EwsUrl
Exchange web services url to connect to.

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

### -EWSTracing
Enable EWS tracing.

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

### -IgnoreSSLCertificate
Ignore SSL validation checks.

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

