function New-GraphBatchRequest
{
    <#
    .SYNOPSIS
    Creates a Microsoft Graph batch request item.

    .DESCRIPTION
    Builds a normalized request object suitable for Invoke-GraphBatch and Graph /$batch payloads.

    .PARAMETER Method
    HTTP method for the subrequest. Allowed values: GET, POST, PUT, PATCH, DELETE.

    .PARAMETER Url
    Relative Graph URL for the subrequest.
    Example: '/me' or '/users?$top=5'.

    .PARAMETER Id
    Request identifier. This value is required by Graph batch requests.

    .PARAMETER Headers
    Optional headers for this subrequest.
    When Body is provided and Content-Type is not set, application/json is used.

    .PARAMETER Body
    Optional body for POST, PUT, or PATCH requests.

    .PARAMETER DependsOn
    Optional list of request IDs this request depends on.

    .INPUTS
    None
    This command does not accept pipeline input.

    .OUTPUTS
    System.Management.Automation.PSCustomObject
    Returns a batch request object.

    .EXAMPLE
    New-GraphBatchRequest -Method GET -Url '/me' -Id '1'

    Creates a batch request item that gets the signed-in user profile.

    .EXAMPLE
    New-GraphBatchRequest -Id '2' -Method PATCH -Url '/users/john.doe@contoso.com' -Body @{ jobTitle = 'Principal Engineer' }

    Creates a batch request item that updates a user.

    .NOTES
    Use this command together with Invoke-GraphBatch to send multiple Graph requests in a single round-trip.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')]
        [string]$Method,
        [Parameter(Mandatory)]
        [string]$Url,
        [Parameter(Mandatory)]
        [string]$Id,
        [Parameter()]
        [System.Collections.Hashtable]$Headers,
        [Parameter()]
        [AllowNull()]
        $Body,
        [Parameter()]
        [string[]]$DependsOn
    )

    process
    {
        if ([string]::IsNullOrWhiteSpace($Url))
        {
            throw 'Url cannot be empty.'
        }

        if ($Url.StartsWith('http', [System.StringComparison]::OrdinalIgnoreCase))
        {
            throw 'Url must be a relative Graph path for batch requests (for example /me or /users?$top=5).'
        }

        $request = [ordered]@{
            method = $Method.ToUpperInvariant()
            url = $Url
        }

        if ([string]::IsNullOrWhiteSpace($Id))
        {
            throw 'Id cannot be empty.'
        }
        $request.id = $Id

        $resolvedHeaders = @{}
        if ($null -ne $Headers)
        {
            foreach ($key in $Headers.Keys)
            {
                $resolvedHeaders[$key] = $Headers[$key]
            }
        }

        if ($PSBoundParameters.ContainsKey('Body'))
        {
            $request.body = $Body

            if (-not $resolvedHeaders.ContainsKey('Content-Type'))
            {
                $resolvedHeaders['Content-Type'] = 'application/json'
            }
        }

        if ($resolvedHeaders.Count -gt 0)
        {
            $request.headers = $resolvedHeaders
        }

        if ($null -ne $DependsOn -and $DependsOn.Count -gt 0)
        {
            $request.dependsOn = $DependsOn
        }

        $requestObject = [PSCustomObject]$request
        $requestObject.PSTypeNames.Insert(0, 'GraphApiHelper.GraphBatchRequest')
        $requestObject
    }
}