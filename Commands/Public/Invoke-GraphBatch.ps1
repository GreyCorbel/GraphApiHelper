function Invoke-GraphBatch
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    <#
    .SYNOPSIS
    Sends a Microsoft Graph batch request.

    .DESCRIPTION
    Collects one or more batch request definitions, builds the Graph $batch payload,
    sends it through Invoke-GraphWithRetry, and returns batch response items.

    .PARAMETER BatchRequest
    One or more Graph batch request objects created by New-GraphBatchRequest.

    .PARAMETER RequestHeaders
    Additional HTTP headers for the outer $batch request.

    .PARAMETER RetryableErrorCodes
    HTTP status codes that should be treated as transient and retried by Invoke-GraphWithRetry.
    Default is 429.

    .PARAMETER OperationName
    The operation name to use for Application Insights logging. Default is 'Invoke-GraphBatch'.

    .OUTPUTS
    System.Object[]
    Returns response items from the Graph batch response.

    .EXAMPLE
    $requests = @(
        New-GraphBatchRequest -Id '1' -Method GET -Url '/me'
        New-GraphBatchRequest -Id '2' -Method GET -Url (New-GraphUri -Uri '/users' -Top 5 -Relative)
        New-GraphBatchRequest -Id '3' -Method POST -Url '/groups' -Body @{ displayName = 'Batch Group'; mailEnabled = $false; mailNickname = 'batch-group'; securityEnabled = $true }
    )

    Invoke-GraphBatch -BatchRequest $requests

    Sends three Graph API requests in one batch and returns the response items. Use New-GraphUri with -Relative to build query strings cleanly.

    .INPUTS
    System.Management.Automation.PSCustomObject[]
    Accepts batch request objects from the pipeline.

    .EXAMPLE
    @(
        New-GraphBatchRequest -Id '1' -Method GET -Url '/me'
        New-GraphBatchRequest -Id '2' -Method GET -Url '/organization'
    ) | Invoke-GraphBatch

    Sends request definitions from the pipeline.

    .EXAMPLE
    Invoke-GraphBatch -BatchRequest $requests -RetryableErrorCodes 429,503

    Sends batch requests while treating 429 and 503 responses from the outer batch call as retryable transient failures.

    .NOTES
    - Uses Invoke-GraphWithRetry internally for reliability.
    - Sends to the /$batch endpoint under the configured BaseUri.
    - Microsoft Graph batch requests support up to 20 subrequests per batch.
    #>
    param
    (
        [Parameter(Mandatory, ValueFromPipeline)]
        [Alias('Requests')]
        [PSCustomObject[]]$BatchRequest,
        [Parameter()]
        [int[]]$RetryableErrorCodes = @(429),
        [Parameter()]
        [System.Collections.Hashtable]$RequestHeaders = @{},
        [Parameter()]
        [string]$OperationName = 'Invoke-GraphBatch'
    )

    begin
    {
        $requests = [System.Collections.Generic.List[hashtable]]::new()
        $ids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    process
    {
        foreach ($item in $BatchRequest)
        {
            if ($null -eq $item)
            {
                continue
            }

            $propertyNames = $item.PSObject.Properties.Name
            if ('id' -notin $propertyNames -or 'method' -notin $propertyNames -or 'url' -notin $propertyNames)
            {
                throw 'Each batch request must include id, method, and url properties. Use New-GraphBatchRequest to create requests.'
            }

            $id = [string]$item.id
            $method = [string]$item.method
            $url = [string]$item.url

            if ([string]::IsNullOrWhiteSpace($id) -or [string]::IsNullOrWhiteSpace($method) -or [string]::IsNullOrWhiteSpace($url))
            {
                throw 'Each batch request must include non-empty id, method, and url values.'
            }

            if (-not $ids.Add($id))
            {
                throw "Duplicate batch request id '$id' is not allowed."
            }

            $normalizedRequest = [ordered]@{
                id = [string]$id
                method = [string]$method.ToUpperInvariant()
                url = [string]$url
            }

            $headers = @{}
            $providedHeaders = $item.headers
            if ($null -ne $providedHeaders)
            {
                foreach ($key in $providedHeaders.Keys)
                {
                    $headers[$key] = $providedHeaders[$key]
                }
            }

            $bodyWasProvided = $false
            if ('body' -in $propertyNames)
            {
                $normalizedRequest.body = $item.body
                $bodyWasProvided = $true
            }

            if ($bodyWasProvided -and -not $headers.ContainsKey('Content-Type'))
            {
                $headers['Content-Type'] = 'application/json'
            }

            if ($headers.Count -gt 0)
            {
                $normalizedRequest.headers = $headers
            }

            if ('dependsOn' -in $propertyNames -and $null -ne $item.dependsOn -and $item.dependsOn.Count -gt 0)
            {
                $normalizedRequest.dependsOn = $item.dependsOn
            }

            [void]$requests.Add($normalizedRequest)
        }
    }

    end
    {
        if ($requests.Count -eq 0)
        {
            Write-Warning 'No batch requests were provided.'
            return
        }

        if ($requests.Count -gt 20)
        {
            throw "Microsoft Graph batch requests support a maximum of 20 subrequests per batch. Received $($requests.Count)."
        }

        $batchUri = New-GraphUri -Uri '/$batch'
        if (-not $PSCmdlet.ShouldProcess($batchUri, "Post Microsoft Graph batch request with $($requests.Count) subrequests"))
        {
            return
        }

        $payload = @{ requests = $requests }
        $result = Invoke-GraphWithRetry `
            -RequestUri $batchUri `
            -Method Post `
            -Body ($payload | ConvertTo-Json -Depth 20) `
            -ContentType 'application/json' `
            -Headers $RequestHeaders `
            -RetryableErrorCodes $RetryableErrorCodes `
            -OperationName $OperationName `
            -Confirm:$false

        if ($null -ne $result.responses)
        {
            $result.responses
        }
        else
        {
            $result
        }
    }
}