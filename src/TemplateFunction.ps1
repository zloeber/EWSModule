function TemplateFunction {
    <#
    .SYNOPSIS
        Template function to use for new functions to this project.
    .DESCRIPTION
    .PARAMETER EWSService
        Exchange web service connection object to use. The default is using the currently connected session.
    .PARAMETER 

    .EXAMPLE
        PS > 
        PS > 

        Description
        -----------
        TBD

    .NOTES
       Author: Zachary Loeber
       Site: http://www.the-little-things.net/
       Requires: Powershell 3.0
       Version History
       1.0.0 - Initial release
    #>
    [CmdletBinding()]
    param(
        [parameter(Position=0, HelpMessage='Connected EWS object.')]
        $EWSService
    )
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

    # Start function here
}