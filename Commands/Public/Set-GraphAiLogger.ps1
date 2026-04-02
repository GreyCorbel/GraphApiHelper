function Set-GraphAiLogger
{
    <#
    .SYNOPSIS
    Sets the Application Insights logger for telemetry
    
    .DESCRIPTION
    Configures the Application Insights logger instance to be used for logging telemetry data during Graph API operations.
    
    .PARAMETER Logger
    The AILogger instance to use for logging. This should be created using the ApplicationInsights module.
    
    .EXAMPLE
    $aiLogger = New-AiLogger -InstrumentationKey 'your-instrumentation-key'
    Set-GraphAiLogger -Logger $aiLogger
    
    Configures the module to use the specified Application Insights logger for telemetry.
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
