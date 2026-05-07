function New-GraphUri
{
    <#
    .SYNOPSIS
    Builds a Microsoft Graph request URL.

    .DESCRIPTION
    Returns a Microsoft Graph request URL using the same query option parameters as Get-GraphData,
    without sending a request.

    The command can build either absolute URLs (using the configured BaseUri) or relative paths
    (for example for Graph batch subrequests). Query options are appended to existing query strings
    and $search values are normalized for Graph search syntax.

    .PARAMETER Uri
    The base Microsoft Graph request URL or relative path.
    When an absolute URL is provided with -Relative, only PathAndQuery is returned.

    .PARAMETER WithSelect
    Optional value for the $select query option.

    .PARAMETER WithFilter
    Optional value for the $filter query option.

    .PARAMETER WithCount
    Adds $count=true to the request.

    .PARAMETER WithExpand
    Optional value for the $expand query option.

    .PARAMETER WithSearch
    Optional value for the $search query option.
    If the value does not start with '(' or '"', the value is automatically wrapped in double quotes.

    .PARAMETER Top
    Optional value for the $top query option.

    .PARAMETER Skip
    Optional value for the $skip query option.

    .PARAMETER Relative
    Returns a relative Graph path instead of prepending the configured BaseUri.
    Use this when building batch request URLs.
    If Uri is absolute, the host and scheme are removed.

    .INPUTS
    None
    This command does not accept pipeline input.

    .OUTPUTS
    System.String
    Returns the fully constructed request URL or relative Graph path.

    .EXAMPLE
    New-GraphUri -Uri '/users' -Top 25

    Returns an absolute Graph URL for /users with the $top query option.

    .EXAMPLE
    New-GraphUri -Uri '/users' -WithSelect 'id,displayName' -WithFilter "accountEnabled eq true"

    Returns a URL with $select and $filter query options.

    .EXAMPLE
    New-GraphUri -Uri '/users' -WithSearch '"displayName:alex"' -WithCount

    Returns a URL with $search and $count query options.

    .EXAMPLE
    New-GraphUri -Uri '/users' -WithSearch 'displayName:alex'

    Returns a URL where the search value is automatically quoted.

    .EXAMPLE
    New-GraphUri -Uri 'https://graph.microsoft.com/v1.0/users' -Relative

    Returns the relative path '/v1.0/users'.

    .EXAMPLE
    New-GraphUri -Uri '/users?$orderby=displayName' -Top 10

    Returns a URL that preserves existing query parameters and appends new ones using '&'.

    .EXAMPLE
    New-GraphUri -Uri '/users' -Top 5 -Relative

    Returns a relative path intended for batch subrequest URLs.

    .NOTES
    If a relative Uri is provided and -Relative is not used, Set-GraphBaseUri must be configured first.
    Values are not URL-encoded by this function. Pass already encoded values when required.

    .LINK
    https://learn.microsoft.com/graph/query-parameters
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [string]$Uri,
        [Parameter()]
        [string]$WithSelect,
        [Parameter()]
        [string]$WithFilter,
        [Parameter()]
        [switch]$WithCount,
        [Parameter()]
        [string]$WithExpand,
        [Parameter()]
        [string]$WithSearch,
        [Parameter()]
        [Nullable[int]]$Top,
        [Parameter()]
        [Nullable[int]]$Skip,
        [Parameter()]
        [switch]$Relative
    )

    process
    {
        if ($Uri.StartsWith('http'))
        {
            if ($Relative)
            {
                # Extract relative path from absolute URI
                $parsedUri = [System.Uri]::new($Uri)
                $Uri = $parsedUri.PathAndQuery
                if ([string]::IsNullOrEmpty($Uri))
                {
                    $Uri = '/'
                }
            }
            # else: Uri is already absolute, use as-is
        }
        else
        {
            # Uri is relative
            if (-not $Relative)
            {
                # Prepend BaseUri
                if(-not $script:graphConnection.BaseUri)
                {
                    throw "BaseUri is not set. Please call Set-GraphBaseUri first or provide a full Uri"
                }
                $Uri = "$($script:graphConnection.BaseUri.TrimEnd('/'))/$($Uri.TrimStart('/'))"
            }
            # else: Uri is already relative, use as-is
        }

        $queryParams = [System.Collections.Generic.List[string]]::new()
        if(-not [string]::IsNullOrWhiteSpace($WithSelect))
        {
            $queryParams.Add("`$select=$($WithSelect.Trim())")
        }
        if(-not [string]::IsNullOrWhiteSpace($WithFilter))
        {
            $queryParams.Add("`$filter=$($WithFilter.Trim())")
        }
        if($WithCount)
        {
            $queryParams.Add('$count=true')
        }
        if(-not [string]::IsNullOrWhiteSpace($WithExpand))
        {
            $queryParams.Add("`$expand=$($WithExpand.Trim())")
        }
        if(-not [string]::IsNullOrWhiteSpace($WithSearch))
        {
            $clause = $WithSearch.Trim()
            if(-not ($clause.StartsWith('(')) -and -not ($clause.StartsWith('"')))
            {
                $clause = "`"$clause`""
            }
            $queryParams.Add("`$search=$clause")
        }
        if($null -ne $Top)
        {
            $queryParams.Add("`$top=$Top")
        }
        if($null -ne $Skip)
        {
            $queryParams.Add("`$skip=$Skip")
        }

        if($queryParams.Count -gt 0)
        {
            $separator = if($Uri.Contains('?')) { '&' } else { '?' }
            $Uri = $Uri + $separator + ($queryParams -join '&')
        }

        return $Uri
    }
}