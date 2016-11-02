---
external help file: EWSModule-help.xml
online version: http://msdn.microsoft.com/en-us/library/dd633699%28v=EXCHG.80%29.aspx
schema: 2.0.0
---

# Test-EWSAutodiscover

## SYNOPSIS
This function uses the EWS Managed API to test the Exchange Autodiscover service.

## SYNTAX

```
Test-EWSAutodiscover [-EmailAddress] <String> [[-Location] <String>] [[-Credential] <PSCredential>]
 [-TraceEnabled] [[-Url] <String>]
```

## DESCRIPTION
This function will retreive the Client Access Server URLs for a specified email address
by querying the autodiscover service of the Exchange server.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Test-Autodiscover -EmailAddress administrator@uclabs.ms -Location internal
```

This example shows how to retrieve the internal autodiscover settings for a user.

### -------------------------- EXAMPLE 2 --------------------------
```
Test-Autodiscover -EmailAddress administrator@uclabs.ms -Credential $cred
```

This example shows how to retrieve the external autodiscover settings for a user.
You can
provide credentials if you do not want to use the Windows credentials of the user calling
the function.

## PARAMETERS

### -EmailAddress
Specifies the email address for the mailbox that should be tested.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Location
Set to External by default, but can also be set to Internal.
This parameter controls whether
the internal or external URLs are returned.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 2
Default value: External
Accept pipeline input: False
Accept wildcard characters: False
```

### -Credential
Specifies a user account that has permission to perform this action.
Type a user name, such as 
"User01" or "Domain01\User01", or enter a PSCredential object, such as one from the Get-Credential cmdlet.

```yaml
Type: PSCredential
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TraceEnabled
Use this switch parameter to enable tracing.
This is used for debugging the XML response from the server.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: 4
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Url
You can use this parameter to manually specifiy the autodiscover url.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 5
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

[http://msdn.microsoft.com/en-us/library/dd633699%28v=EXCHG.80%29.aspx](http://msdn.microsoft.com/en-us/library/dd633699%28v=EXCHG.80%29.aspx)

[http://www.the-little-things.net/](http://www.the-little-things.net/)

[https://www.github.com/zloeber/EWSModule](https://www.github.com/zloeber/EWSModule)

