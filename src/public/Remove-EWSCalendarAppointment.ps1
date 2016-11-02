function Remove-EWSCalendarAppointment {
    <#
    .SYNOPSIS
    Remove a calendar appointment object from a mailbox.
    .DESCRIPTION
    Remove a calendar appointment object from a mailbox.
    .PARAMETER EWSService
    Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER Appointment
    EWS Calendar appointment object
    .PARAMETER DeleteMode
    Method of deletion for the appointment. Can be 'HardDelete','SoftDelete', or 'MoveToDeletedItems'
    .PARAMETER CancellationMode
    How cancellation notices will be sent upon deletion. Can be 'SendToNone','SendOnlyToAll','SendOnlyToChanged','SendToAllAndSaveCopy', or 'SendToChangedAndSaveCopy'

    .EXAMPLE
    PS > Remove-EWSCalendarAppointment -Appointment $Appt -DeleteMode 'SoftDelete' -CancellationMode 'SendToAllAndSaveCopy'

    Removes the calendar appointment stored in $Appt using softdelete and sending a cancellation notice and saving a copy of the sent cancellation. 

    .LINK
    http://www.the-little-things.net/

    .LINK
    https://www.github.com/zloeber/EWSModule

    .NOTES
    Author: Zachary Loeber
    Requires: Powershell 3.0
    Version History
    1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService,
        [parameter(Position=1, Mandatory=$true,ValueFromPipeline=$true, HelpMessage='Calendar appointment object to remove.')]
        [Microsoft.Exchange.WebServices.Data.Appointment]$Appointment,
        [parameter(Position=2, HelpMessage='Deletion mode.')]
        [ValidateSet('HardDelete','SoftDelete','MoveToDeletedItems')]
        [string]$DeleteMode = 'HardDelete',
        [parameter(Position=3, HelpMessage='Cancellation mode.')]
        [ValidateSet('SendToNone','SendOnlyToAll','SendOnlyToChanged','SendToAllAndSaveCopy','SendToChangedAndSaveCopy')]
        [string]$CancellationMode = 'SendToNone'
    )
   Begin {
        # Pull in all the caller verbose,debug,info,warn and other preferences
        Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
        $FunctionName = $MyInvocation.MyCommand
        
        if (-not (Get-EWSModuleInitializationState)) {
            throw "$($FunctionName): EWS Module has not been initialized. Try running Initialize-EWS to rectify."
        }
        
        if ($EWSService -eq $null) {
            Write-Verbose "$($FunctionName): Using module local ews service object"
            $EWSService = Get-EWSService
        }
        
        if ($EWSService -eq $null) {
            throw "$($FunctionName): EWS connection has not been established. Create a new connection with Connect-EWS first."
        }
   }
    Process {
         $Appointment | Foreach {
            Write-Verbose "$($FunctionName): Deleting existing appointment with the subject of $($_.Subject)"
            try {                
                $_.Delete([ews_deletemode]::$DeleteMode,[ews_sendcancellationmode]::$CancellationMode)
            }
            catch {
                Write-Warning "$($FunctionName): Unable to DELETE existing appointment!"
            }
         }
    }
    End {}
}