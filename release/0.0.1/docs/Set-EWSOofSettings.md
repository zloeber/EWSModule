---
external help file: EWSModule-help.xml
online version: http://www.the-little-things.net/
schema: 2.0.0
---

# Set-EWSOofSettings

## SYNOPSIS
Set the out of office settings for a mailbox.

## SYNTAX

### Default (Default)
```
Set-EWSOofSettings [[-EWSService] <ExchangeService>] [-Mailbox] <String> [[-State] <String>]
```

### Scheduled
```
Set-EWSOofSettings [[-EWSService] <ExchangeService>] [-Mailbox] <String> [[-State] <String>]
 [[-ExternalAudience] <String>] [[-StartTime] <DateTime>] [[-EndTime] <DateTime>] [[-InternalReply] <String>]
 [[-ExternalReply] <String>]
```

### Disabled
```
Set-EWSOofSettings [[-EWSService] <ExchangeService>] [-Mailbox] <String> [[-State] <String>]
```

### Enabled
```
Set-EWSOofSettings [[-EWSService] <ExchangeService>] [-Mailbox] <String> [[-State] <String>]
 [[-ExternalAudience] <String>] [[-StartTime] <DateTime>] [[-EndTime] <DateTime>] [[-InternalReply] <String>]
 [[-ExternalReply] <String>]
```

## DESCRIPTION
Set the out of office settings for a mailbox.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Set-EWSOOFSettings -Mailbox mailbox@domain.com
```

Description
--------------
Disables the OOF settings for mailbox@domain.com

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
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Mailbox
Mailbox to target.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -State
State of OOF for the mailbox.
Can be Enabled, Disabled, or Scheduled.
Defaults to Disabled.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: Disabled
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalAudience
Whom will get OOF externally.
Can be All, Known, or None.
Defaults to All.

```yaml
Type: String
Parameter Sets: Scheduled, Enabled
Aliases: 

Required: False
Position: 4
Default value: All
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartTime
Start time that OOF replies will be scheduled.

```yaml
Type: DateTime
Parameter Sets: Scheduled, Enabled
Aliases: 

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EndTime
End time that OOF replies will be enabled or scheduled.

```yaml
Type: DateTime
Parameter Sets: Scheduled, Enabled
Aliases: 

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InternalReply
Internal OOF message.

```yaml
Type: String
Parameter Sets: Scheduled, Enabled
Aliases: 

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExternalReply
External OOF message.

```yaml
Type: String
Parameter Sets: Scheduled, Enabled
Aliases: 

Required: False
Position: 8
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

