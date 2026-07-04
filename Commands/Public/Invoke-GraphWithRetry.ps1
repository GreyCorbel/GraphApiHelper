function Invoke-GraphWithRetry
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    <#
    .SYNOPSIS
    Invokes a Graph API with automatic retry logic for throttling
    
    .DESCRIPTION
    Executes a Microsoft Graph API request with built-in retry logic to handle transient HTTP responses.
    The function retries retryable status codes (429 by default), using Retry-After when provided,
    or incremental backoff based on DefaultBackOffSeconds and the retry attempt number.
    If the request returns paged results, it retrieves only a single page - callers should use Get-GraphData for automatic pagination.
    
    Supports Application Insights logging when an AILogger instance is provided when importing the module
    
    .PARAMETER RequestUri
    The complete Microsoft Graph API request URL including query parameters.
    Example: 'https://graph.microsoft.com/v1.0/users?$top=10'
    
    .PARAMETER Method
    The HTTP method to use for the request. Valid values are: Get, Post, Put, Patch, Delete.
    Default is 'Get'.
    
    .PARAMETER Body
    The request body for Post, Put, or Patch requests. Can be a string or object that will be sent with the request.
    
    .PARAMETER ContentType
    The content type for the request body. Default is 'application/json'.
    
    .PARAMETER Headers
    Additional HTTP headers to include in the request. The Authorization header will be automatically added.
    
    .PARAMETER OperationName
    The operation name to use for Application Insights logging. Default is 'Invoke-GraphWithRetry'.

    .PARAMETER RetryableErrorCodes
    HTTP status codes that should trigger retries. Default is 429.

    .PARAMETER MaxRetries
    Maximum retry threshold used by the retry loop before the error is written. Default is 100.

    .PARAMETER DefaultBackOffSeconds
    Fallback delay in seconds used when the response does not include Retry-After.

    .PARAMETER AuthorizationHeader
    Optional pre-obtained authorization header hashtable (e.g. from Get-GraphAuthorizationHeader).
    When provided, token acquisition is skipped and this header is used directly.
    Useful for reusing a token across multiple calls or for testing.

    .INPUTS
    None
    This command does not accept pipeline input.
    
    .OUTPUTS
    System.Object
    Returns the response from the Graph API call.
    
    .EXAMPLE
    Invoke-GraphWithRetry -RequestUri 'https://graph.microsoft.com/v1.0/users'
    
    Retrieves users from Microsoft Graph using the default GET method.
    
    .EXAMPLE
    $body = @{ displayName = 'Test Group' } | ConvertTo-Json
    Invoke-GraphWithRetry -RequestUri 'https://graph.microsoft.com/v1.0/groups' -Method Post -Body $body
    
    Creates a new group in Microsoft Graph.
    
    .EXAMPLE
    Invoke-GraphWithRetry -RequestUri 'https://graph.microsoft.com/v1.0/users/user@domain.com' -Method Delete
    
    Deletes a user from Microsoft Graph.

    .EXAMPLE
    Invoke-GraphWithRetry -RequestUri 'https://graph.microsoft.com/v1.0/users/user@domain.com' -Method Delete -WhatIf

    Shows what delete request would run without calling Microsoft Graph.
    
    .NOTES
    - Automatically retries status codes listed in RetryableErrorCodes (429 by default)
    - Uses Retry-After for 429 responses when available; otherwise uses incremental backoff
    - Uses the authentication factory configured via Set-GraphAadFactory
    - Uses the Graph scopes configured via Set-GraphScopes
    - Supports Application Insights telemetry when configured
    - Supports -WhatIf and -Confirm via ShouldProcess
    #>
    param
    (
        [Parameter(Mandatory)]
        [Alias('Uri')]
        [string]$RequestUri,
        [Parameter()]
        [ValidateSet('Get', 'Post', 'Put', 'Patch', 'Delete')]
        $method = 'Get',
        [Parameter()]
        $Body,
        [Parameter()]
        $ContentType = 'application/json',
        [parameter()]
        [System.Collections.Hashtable]
        $Headers = @{},
        [Parameter()]
        $OperationName = 'Invoke-GraphWithRetry',
        [Parameter()]
        [int[]]$RetryableErrorCodes = @(429),
        [Parameter()]
        [int]$MaxRetries = 100,
        [Parameter()]
        [int]$DefaultBackOffSeconds = 1,
        [Parameter()]
        [hashtable]$AuthorizationHeader
    )

    begin
    {
        $retries = 0
        $graphUri = New-GraphUri -Uri $RequestUri
    }
    process
    {
        if (-not $PSCmdlet.ShouldProcess($graphUri, "$method Microsoft Graph request"))
        {
            return
        }

        do
        {
            if($null -eq $AuthorizationHeader)
            {
                $AuthorizationHeader = Get-GraphAuthorizationHeader
            }
            Write-Verbose "Invoking Graph API: $graphUri with method $method. Attempt #$($retries + 1)"
            $headers['Authorization'] = $AuthorizationHeader['Authorization']
            $resultCode = 'Ok'
            try {
                $requestStart = [DateTime]::UtcNow

                switch($method)
               {
                    {$_ -in @('Get', 'Delete')} {
                        $result = Invoke-RestMethod -method $method -Uri $graphUri -headers $headers -ErrorAction Stop -Verbose:$VerbosePreference
                        break;
                    }
                    {$_ -in @('Post', 'Patch', 'Put')} {
                        $result = Invoke-RestMethod -method $method -Uri $graphUri -body $body -headers $headers -ContentType $contentType -ErrorAction Stop -Verbose:$VerbosePreference
                        break;
                    }
                }
                if($script:graphConnection.AiLogger)
                {
                    Write-AiDependency -Target 'graph.microsoft.com' -DependencyType 'Graph API' -Name $OperationName -Data $graphUri -Start $requestStart -ResultCode 'Ok' -Success $true -Connection $script:graphConnection.AiLogger
                }
                $result
                break;  #do-while
            }
            catch {
                $err = $_
                if($null -ne $script:graphConnection.AiLogger)
                {
                    Write-AiException -Exception $err.Exception -Connection $script:graphConnection.AiLogger
                }
                if($null -ne $err.exception.Response.StatusCode)
                {
                    $resultCode = $err.exception.Response.StatusCode
                }
                else {
                    $resultCode = 'Unknown'
                }

                if($retries -le $MaxRetries -and ($err.exception.Response.StatusCode -in $RetryableErrorCodes))
                {
                    $retries++
                    switch($err.exception.Response.StatusCode)
                    {
                        429 {
                            $retryAfter = $err.exception.Response.Headers['Retry-After']
                            if($null -eq $retryAfter)
                            {
                                $retryAfter = $DefaultBackOffSeconds * $retries
                            }
                            $waitTime = [int]$retryAfter
                            break;
                        }
                        default {
                            $waitTime = $DefaultBackOffSeconds * $retries
                            break;
                        }
                    }
                    Write-Warning "Retrying because of status code $($err.exception.Response.StatusCode) for $waitTime secs"
                    start-sleep -Seconds $waitTime
                }
                else {
                    Write-Error -ErrorRecord $_
                    break;
                }
            }
            finally
            {
                if($null -ne $script:graphConnection.AiLogger)
                {
                    Write-AiDependency -Target 'graph.microsoft.com' -DependencyType 'Graph API' -Name $OperationName -Data $graphUri -Start $requestStart -ResultCode $resultCode -Success ($resultCode -eq 'Ok') -Connection $script:graphConnection.AiLogger
                }
            }
        } while($true)
    }
}
