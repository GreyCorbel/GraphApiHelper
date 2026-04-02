function Set-GraphBaseUri
{
    <#
    .SYNOPSIS
    Sets the base URI used for Microsoft Graph API requests.

    .DESCRIPTION
    Configures the base URI used to build absolute request URIs when a relative path is supplied
    to commands such as Invoke-GraphWithRetry and Get-GraphData.

    .PARAMETER BaseUri
    The base URI to use for Graph requests. Defaults to https://graph.microsoft.com/v1.0 when
    the module is imported.

    .EXAMPLE
    Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/v1.0'

    Uses the global Microsoft Graph endpoint.

    .EXAMPLE
    Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'

    Uses the Microsoft Graph US Government endpoint.
    #>
    param
    (
        [Parameter(Mandatory)]
        [string]$BaseUri
    )

    process
    {
        $script:graphConnection.BaseUri = $BaseUri
    }
}
