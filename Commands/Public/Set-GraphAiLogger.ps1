function Set-GraphAiLogger
{
    <#
    .SYNOPSIS
    Sets the Application Insights logger for telemetry
    
    .DESCRIPTION
    Configures the Application Insights logger instance to be used for logging telemetry data during Graph API operations.
    
    .PARAMETER Logger
    The AILogger instance to use for logging. This should be created using the ApplicationInsights module.

    .INPUTS
    None
    This command does not accept pipeline input.

    .OUTPUTS
    None
    This command updates module configuration and does not return an object.
    
    .EXAMPLE
    $logger = Connect-AiLogger -ConnectionString 'InstrumentationKey=...'
    Set-GraphAiLogger -Logger $logger
    
    Configures the module to use the specified Application Insights logger for telemetry.

    .NOTES
    Invoke-GraphWithRetry uses this logger for dependency and exception telemetry when configured.
    #>
    param
    (
        [Parameter(Mandatory)]
        $Logger
    )

    process
    {
        $script:graphConnection.AiLogger = $Logger
    }
}
