function Set-EWSModuleInitializationState ([bool]$State) {
    $script:EWSModuleInitialized = $State
}