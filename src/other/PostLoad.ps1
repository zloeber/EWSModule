
#region Module Cleanup
$ExecutionContext.SessionState.Module.OnRemove = {
        if ( Initialize-EWS -Uninitialize ) {}
        else { Write-Warning "Unable to uninitialize module" }
}

$null = Register-EngineEvent -SourceIdentifier ( [System.Management.Automation.PsEngineEvent]::Exiting ) -Action {
        if ( Initialize-EWS -Uninitialize ) {}
        else { Write-Warning "Unable to uninitialize module" }
}
#endregion Module Cleanup