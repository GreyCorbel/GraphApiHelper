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

    .INPUTS
    None
    This command does not accept pipeline input.

    .OUTPUTS
    None
    This command updates module configuration and does not return an object.

    .EXAMPLE
    Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/v1.0'

    Uses the global Microsoft Graph endpoint.

    .EXAMPLE
    Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'

    Uses the Microsoft Graph US Government endpoint.

    .NOTES
    This value is used when commands receive relative Uri values, for example in New-GraphUri and Invoke-GraphWithRetry.
    #>
    param
    (
        [Parameter(Mandatory)]
        [string]$BaseUri
    )

    process
    {
        $uri = New-Object System.Uri($BaseUri.Trim().TrimEnd('/'))
        if($uri.Segments.Length -lt 2)
        {
            throw "Invalid BaseUri. Please provide a valid URI with at least one segment."
        }
        if($uri.Segments[1].TrimEnd('/') -notin @('v1.0', 'beta'))
        {
            throw "BaseUri must include a version segment (e.g. 'v1.0' or 'beta')."
        }
        $script:graphConnection.BaseUri = $uri
    }
}
