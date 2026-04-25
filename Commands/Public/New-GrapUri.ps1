function New-GrapUri
{
    <#
    .SYNOPSIS
    Builds a Microsoft Graph request URL.

    .DESCRIPTION
    Returns a Microsoft Graph request URL using the same query option parameters as Get-GraphData,
    without sending a request.

    .PARAMETER Uri
    The base Microsoft Graph request URL or relative path.

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

    .PARAMETER Top
    Optional value for the $top query option.

    .PARAMETER Skip
    Optional value for the $skip query option.

    .PARAMETER Relative
    Returns a relative Graph path instead of prepending the configured BaseUri.
    Use this when building batch request URLs.

    .OUTPUTS
    System.String
    Returns the fully constructed request URL.
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
        if(-not [string]::IsNullOrEmpty($WithSelect))
        {
            $queryParams.Add("`$select=$WithSelect")
        }
        if(-not [string]::IsNullOrEmpty($WithFilter))
        {
            $queryParams.Add("`$filter=$WithFilter")
        }
        if($WithCount)
        {
            $queryParams.Add('$count=true')
        }
        if(-not [string]::IsNullOrEmpty($WithExpand))
        {
            $queryParams.Add("`$expand=$WithExpand")
        }
        if(-not [string]::IsNullOrEmpty($WithSearch))
        {
            $queryParams.Add("`$search=`"$WithSearch`"")
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