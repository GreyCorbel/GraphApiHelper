function Invoke-GraphSearchQuery
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    <#
    .SYNOPSIS
    Queries the Microsoft Graph Search API (/search/query) with automatic result-page traversal.

    .DESCRIPTION
    Executes a Microsoft Graph Search query against the /search/query endpoint for a single entity type
    and automatically pages through results using the from/size offset model used by the Search API
    (Search does not use @odata.nextLink; instead it reports moreResultsAvailable and expects the caller
    to resubmit the query with an incremented from value).

    Use Invoke-GraphSearchQuery for content search across mail, drive items, sites, lists, people, Teams
    messages, and Search-connector external items.

    .PARAMETER EntityType
    The Microsoft Graph Search entity type to query. Valid values are: list, listItem, site, drive, driveItem,
    message, event, person, chatMessage, externalItem, acronym, bookmark, qna.

    .PARAMETER QueryString
    The search query string (KQL-like syntax supported by Microsoft Search).

    .PARAMETER Fields
    Optional list of properties to return for each hit (the 'fields' request property). When omitted, Graph
    returns its default field set for the entity type.

    .PARAMETER SortProperties
    Optional array of sort specifications. Each entry is a hashtable with 'Name' and optionally 'IsDescending',
    for example @{ Name = 'lastModifiedDateTime'; IsDescending = $true }.

    .PARAMETER ContentSources
    Optional list of external connection content source identifiers. Only applies when -EntityType is externalItem.

    .PARAMETER Region
    Geographic region required by application (non-delegated) permission searches for some entity types.
    See Microsoft Graph Search documentation for the required value in your tenant.

    .PARAMETER EnableTopResults
    Requests relevance-ranked top results ordering when supported by the entity type.

    .PARAMETER Size
    Number of hits requested per page. Default is 25. Microsoft Graph enforces an entity-type-specific maximum.

    .PARAMETER From
    Zero-based offset of the first hit to return. Use this to resume a previously interrupted search
    (see -ResponseMetadataVariable). Default is 0.

    .PARAMETER NoContinue
    When specified, retrieves only the first page of results and does not continue fetching subsequent pages.

    .PARAMETER RetryableErrorCodes
    HTTP status codes that should be treated as transient and retried by Invoke-GraphWithRetry. Default is 429.

    .PARAMETER OperationName
    The operation name to use for Application Insights logging. Default is 'Invoke-GraphSearchQuery'.

    .PARAMETER AuthorizationHeader
    Optional pre-obtained authorization header hashtable (e.g. from Get-GraphAuthorizationHeader).
    When provided, token acquisition is skipped and this header is used directly.

    .PARAMETER ResponseMetadataVariable
    Name of a variable in the caller's scope that receives search response metadata (Total, MoreResultsAvailable,
    NextFrom, Size) from the last page returned. The variable is also populated when result traversal is
    interrupted (e.g. Ctrl+C) or stops early due to a non-terminating error, so NextFrom can be passed back in
    as -From to resume the search from that point instead of starting over.

    .INPUTS
    None
    This command does not accept pipeline input.

    .OUTPUTS
    System.Object
    Returns Search hit objects, automatically requesting subsequent pages until moreResultsAvailable is false.

    .EXAMPLE
    Invoke-GraphSearchQuery -EntityType message -QueryString 'subject:invoice'

    Searches mail messages for the current user context, paging through all matching results.

    .EXAMPLE
    Invoke-GraphSearchQuery -EntityType driveItem -QueryString 'quarterly report' -Fields 'name','webUrl' -Size 50

    Searches drive items, requesting only the name and webUrl fields, 50 hits per page.

    .EXAMPLE
    Invoke-GraphSearchQuery -EntityType driveItem -QueryString 'roadmap' -ResponseMetadataVariable meta -NoContinue
    $meta.NextFrom

    Retrieves only the first page and captures the offset needed to resume with -From on the next call.

    .EXAMPLE
    Invoke-GraphSearchQuery -EntityType externalItem -QueryString 'contoso' -ContentSources 'connections/contoso-kb' -Region 'NAM'

    Searches an external connection's indexed content, specifying the required region for application permissions.

    .EXAMPLE
    Invoke-GraphSearchQuery -EntityType message -QueryString 'subject:invoice' -WhatIf

    Shows what search request would be executed without calling Microsoft Graph.

    .NOTES
    - Queries a single entity type per call; call the command again for additional entity types.
    - Automatically pages using the from/size model instead of @odata.nextLink.
    - Uses Invoke-GraphWithRetry internally for throttling protection.
    - Uses the authentication factory configured via Set-GraphAadFactory.
    - Uses the Graph scopes configured via Set-GraphScopes.
    - Supports -WhatIf and -Confirm via ShouldProcess.

    .LINK
    https://learn.microsoft.com/graph/search-concept-overview
    #>
    param
    (
        [Parameter(Mandatory)]
        [ValidateSet('list', 'listItem', 'site', 'drive', 'driveItem', 'message', 'event', 'person', 'chatMessage', 'externalItem', 'acronym', 'bookmark', 'qna')]
        [string]$EntityType,
        [Parameter(Mandatory)]
        [string]$QueryString,
        [Parameter()]
        [string[]]$Fields,
        [Parameter()]
        [hashtable[]]$SortProperties,
        [Parameter()]
        [string[]]$ContentSources,
        [Parameter()]
        [string]$Region,
        [Parameter()]
        [switch]$EnableTopResults,
        [Parameter()]
        [ValidateRange(1, 1000)]
        [int]$Size = 25,
        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [int]$From = 0,
        [Parameter()]
        [switch]$NoContinue,
        [Parameter()]
        [int[]]$RetryableErrorCodes = @(429),
        [Parameter()]
        [string]$OperationName = 'Invoke-GraphSearchQuery',
        [Parameter()]
        [hashtable]$AuthorizationHeader,
        [Parameter()]
        [string]$ResponseMetadataVariable
    )

    process
    {
        $uri = New-GraphUri -Uri '/search/query'

        if (-not $PSCmdlet.ShouldProcess($uri, "Search $EntityType for '$QueryString'"))
        {
            return
        }

        $offset = $From
        $lastContainer = $null
        try
        {
            do
            {
                $request = [ordered]@{
                    entityTypes = @($EntityType)
                    query       = @{ queryString = $QueryString }
                    from        = $offset
                    size        = $Size
                }
                if($Fields.Count -gt 0)
                {
                    $request.fields = $Fields
                }
                if($SortProperties.Count -gt 0)
                {
                    $request.sortProperties = $SortProperties
                }
                if($ContentSources.Count -gt 0)
                {
                    $request.contentSources = $ContentSources
                }
                if(-not [string]::IsNullOrWhiteSpace($Region))
                {
                    $request.region = $Region
                }
                if($EnableTopResults)
                {
                    $request.enableTopResults = $true
                }

                $payload = @{ requests = @($request) } | ConvertTo-Json -Depth 10

                #get page of results
                $result = Invoke-GraphWithRetry `
                    -RequestUri $uri `
                    -Method Post `
                    -Body $payload `
                    -OperationName $OperationName `
                    -Confirm:$false `
                    -ErrorAction $ErrorActionPreference `
                    -RetryableErrorCodes $RetryableErrorCodes `
                    -AuthorizationHeader $AuthorizationHeader

                if($null -eq $result)
                {
                    #non-terminating error already reported by Invoke-GraphWithRetry; stop paginating
                    break
                }

                $container = $result.value[0].hitsContainers[0]
                $lastContainer = $container

                if($null -ne $container.hits)
                {
                    $container.hits
                }

                $moreAvailable = [bool]$container.moreResultsAvailable
                $offset += $Size
            }
            while($moreAvailable -and -not $NoContinue)
        }
        finally
        {
            #runs on normal completion and on interruption (Ctrl+C);
            #if interrupted mid-pagination, or if a later page fetch failed non-terminally,
            #$lastContainer still holds the last successfully fetched page whose offset
            #lets the caller resume the search from that point via -From
            if(-not [string]::IsNullOrEmpty($ResponseMetadataVariable) -and $null -ne $lastContainer)
            {
                $metadata = [PSCustomObject]@{
                    Total                = $lastContainer.total
                    MoreResultsAvailable = $lastContainer.moreResultsAvailable
                    NextFrom             = $offset
                    Size                 = $Size
                }
                $PSCmdlet.SessionState.PSVariable.Set($ResponseMetadataVariable, $metadata)
            }
        }
    }
}
