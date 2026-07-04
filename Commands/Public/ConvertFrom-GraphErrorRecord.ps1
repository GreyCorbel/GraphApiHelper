function ConvertFrom-GraphErrorRecord
{
    <#
    .SYNOPSIS
    Extracts Microsoft Graph error details from a PowerShell error record.

    .DESCRIPTION
    Parses the ErrorDetails payload from a PowerShell ErrorRecord and returns the
    deserialized Graph error object when it contains an error message.

    This helper is useful when handling failures from Invoke-GraphWithRetry and
    other commands that return Graph error payloads in JSON format.

    .PARAMETER ErrorRecord
    The PowerShell ErrorRecord to parse.

    .INPUTS
    System.Management.Automation.ErrorRecord
    Accepts error records from the pipeline.

    .OUTPUTS
    System.Object
    Returns the deserialized Graph error object when available.

    .EXAMPLE
    try {
        Invoke-GraphWithRetry -RequestUri 'https://graph.microsoft.com/v1.0/users/does-not-exist' -ErrorAction Stop
    }
    catch {
        $_ | ConvertFrom-GraphErrorRecord
    }

    Parses the Graph error payload from the caught exception.

    .EXAMPLE
    $details = $Error[0] | ConvertFrom-GraphErrorRecord

    Parses the most recent error record and returns Graph error details when present.

    .NOTES
    Returns nothing when the error details are not JSON or do not contain error.message.

    .LINK
    https://github.com/GreyCorbel/GraphApiHelper
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    process
    {
        $details = $ErrorRecord.ErrorDetails | ConvertFrom-Json -ErrorAction SilentlyContinue
        if($null -ne $details.error.message)
        {
            $details
        }
    }
}
