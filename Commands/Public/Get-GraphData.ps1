function Get-GraphData
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    <#
    .SYNOPSIS
    Retrieves data from Microsoft Graph API with automatic pagination
    
    .DESCRIPTION
    Executes a Microsoft Graph API GET request and automatically handles pagination by following @odata.nextLink references.
    This function retrieves all pages of data and returns the complete dataset. It uses Invoke-GraphWithRetry internally,
    so it inherits automatic retry logic for throttling.
    
    The function intelligently handles both single objects and arrays of results from the Graph API.
    
    .PARAMETER RequestUri
    The complete Microsoft Graph API request URL including query parameters.
    Example: 'https://graph.microsoft.com/v1.0/users'

    .PARAMETER WithSelect
    Optional value for the $select query option.
    Example: 'id,displayName,userPrincipalName'

    .PARAMETER WithFilter
    Optional value for the $filter query option.
    Example: "accountEnabled eq true"

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

    .PARAMETER RetryableErrorCodes
    HTTP status codes that should be treated as transient and retried by Invoke-GraphWithRetry.
    Default is 429.
    
    .PARAMETER OperationName
    The operation name to use for Application Insights logging. Default is 'Get-GraphData'.

    .PARAMETER AdditionalHeaders
    Additional HTTP headers to include in requests (for example ConsistencyLevel for advanced queries).

    .PARAMETER NoContinue
    When specified, retrieves only the first page and does not follow @odata.nextLink.

    .INPUTS
    None
    This command does not accept pipeline input.
    
    .OUTPUTS
    System.Object[]
    Returns all objects from the Graph API response, automatically handling pagination.
    
    .EXAMPLE
    Get-GraphData -RequestUri 'https://graph.microsoft.com/v1.0/users'
    
    Retrieves all users from Microsoft Graph, automatically paginating through all result pages.
    
    .EXAMPLE
    Get-GraphData -RequestUri 'https://graph.microsoft.com/v1.0/groups?$filter=startswith(displayName,''Sales'')'
    
    Retrieves all groups whose display name starts with 'Sales', handling pagination automatically.
    
    .EXAMPLE
    Get-GraphData -RequestUri 'https://graph.microsoft.com/v1.0/me/messages?$top=50' -OperationName 'GetUserMessages'
    
    Retrieves all messages for the current user with custom operation name for Application Insights tracking.

    .EXAMPLE
    Get-GraphData -RequestUri 'https://graph.microsoft.com/v1.0/users' -WithSelect 'id,displayName' -WithFilter "startswith(displayName,'A')" -WithCount -AdditionalHeaders @{ ConsistencyLevel = 'eventual' }

    Retrieves users with query options built from parameters and the required advanced query header.

    .EXAMPLE
    Get-GraphData -RequestUri 'https://graph.microsoft.com/v1.0/users' -WhatIf

    Shows what request would be executed without calling Microsoft Graph.

    .EXAMPLE
    Get-GraphData -RequestUri 'https://graph.microsoft.com/v1.0/users' -RetryableErrorCodes 429,503

    Retrieves users while treating 429 and 503 responses as retryable transient failures.
    
    .NOTES
    - Automatically handles pagination via @odata.nextLink
    - Uses Invoke-GraphWithRetry internally for throttling protection
    - Suitable for large datasets that span multiple pages
    - Uses the authentication factory configured via Set-GraphAadFactory
    - Uses the Graph scopes configured via Set-GraphScopes
    - Supports -WhatIf and -Confirm via ShouldProcess
    #>
    param
    (
        [Parameter(Mandatory)]
        [Alias('Uri')]
        [string]$RequestUri,
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
        [int[]]$RetryableErrorCodes = @(429),
        [Parameter()]
        [string]$OperationName = 'Get-GraphData',
        [Parameter()]
        [System.Collections.Hashtable]$AdditionalHeaders = @{},
        [Parameter()]
        [switch]$NoContinue
    )

    process
    {
        $uri = New-GraphUri -Uri $RequestUri -WithSelect $WithSelect -WithFilter $WithFilter -WithCount:$WithCount -WithExpand $WithExpand -WithSearch $WithSearch -Top $Top -Skip $Skip

        if (-not $PSCmdlet.ShouldProcess($uri, 'Get Microsoft Graph data with automatic pagination'))
        {
            return
        }

        while($true)
        {
            try {
                #get page of results
                $result = Invoke-GraphWithRetry -RequestUri $uri -method Get -Headers $AdditionalHeaders -OperationName $OperationName -Confirm:$false -ErrorAction Stop -RetryableErrorCodes $RetryableErrorCodes
                if($null -ne $result.value)
                {
                    #returning array of results
                    $result.value
                }
                else
                {
                    #returning single object
                    $result
                }
                $uri = $result.'@odata.nextLink'
                if([string]::IsNullOrEmpty($uri) -or $NoContinue)
                {
                    #no more pages or we just wanted first page
                    break;
                }
            }
            catch {
                $err = $_
                $shouldContinue = $false
                switch($ErrorActionPreference)
                {
                    'Stop' { throw }
                    'Continue' { 
                        Write-Error "Could not retrieve data for uri: $uri. Error: $($err.Exception.Message)"
                        $shouldContinue = $false
                        break 
                    }
                    'SilentlyContinue' { 
                        Write-Warning "Could not retrieve data for uri: $uri. Error: $($err.Exception.Message)"
                        $shouldContinue = $false
                        break 
                    }
                    'Inquire' {
                        $response = $PSCmdlet.ShouldContinue("Error retrieving data for uri: $uri. Do you want to continue?", "Error: $($err.Exception.Message)")
                        if(-not $response)
                        {
                            $shouldContinue = $false
                            break
                        }
                    }
                }
                if(-not $shouldContinue)
                {
                    break
                }
            }
        }
    }
}
