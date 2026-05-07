#region Public commands
function Add-GraphLargeFile
{
    <#
    .SYNOPSIS
    Uploads large files to Microsoft Graph using the resumable upload protocol
    
    .DESCRIPTION
    Uploads large files to Microsoft Graph (OneDrive, SharePoint, etc.) using the resumable upload session API.
    This function handles files of any size by splitting them into chunks and uploading them sequentially.
    
    The upload uses 5MB chunks (320KB * 16) which is optimal for Graph API uploads and supports resumable uploads
    in case of network interruptions. The function automatically creates an upload session and manages the chunked
    upload process.
    
    .PARAMETER LocalFilePath
    The full path to the local file to upload. The file must exist and be readable.
    
    .PARAMETER GraphFilePath
    The Microsoft Graph API path where the file should be uploaded, excluding the ':/createUploadSession' suffix.
    Example: 'https://graph.microsoft.com/v1.0/me/drive/root:/Documents/myfile.pdf'
    
    .EXAMPLE
    Add-GraphLargeFile -LocalFilePath 'C:\Files\presentation.pptx' -GraphFilePath 'https://graph.microsoft.com/v1.0/me/drive/root:/Documents/presentation.pptx'
    
    Uploads a PowerPoint file to the current user's OneDrive Documents folder.
    
    .EXAMPLE
    Add-GraphLargeFile -LocalFilePath 'C:\Videos\training.mp4' -GraphFilePath 'https://graph.microsoft.com/v1.0/sites/{site-id}/drive/root:/Videos/training.mp4' -Verbose
    
    Uploads a video file to a SharePoint site's Videos folder with verbose output showing upload progress.

    .OUTPUTS
    None
    The function streams upload chunk requests and does not emit the final driveItem object.

    .INPUTS
    None
    This command does not accept pipeline input.
    
    .NOTES
    - Uses 5MB chunks for optimal performance
    - Automatically handles upload session creation
    - Supports conflict behavior of 'replace' - existing files will be overwritten
    - Uses Invoke-GraphWithRetry internally for reliability
    - Enable -Verbose to see detailed upload progress
    - Uses the authentication factory configured via Set-GraphAadFactory
    
    .LINK
    https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $LocalFilePath,
        [Parameter(Mandatory)]
        $GraphFilePath
    )

    begin
    {
        $chunkSize = 320KB * 16 # 5MB chunks
        $graphUri = New-GraphUri -Uri "$GraphFilePath"
    }
    process
    {
        $item = Get-Item -Path $LocalFilePath
        $fileSize = $item.length
        $fileStream = [System.IO.File]::OpenRead($item.FullName)
        Write-Verbose "Filesize: $fileSize"
        Write-Verbose "Chunksize: $chunkSize"
        try {
            $payload =  @{
                item = @{
                    '@microsoft.graph.conflictBehavior' = 'replace' 
                }
            }
            Write-Verbose "Requesting upload session on $graphUri`:/createUploadSession"
            $uploadSession = Invoke-GraphWithRetry `
                -RequestUri "$graphUri`:/createUploadSession" `
                -method Post `
                -body ($payload | ConvertTo-Json -Depth 10) `
                -ContentType 'application/json' `
                -ErrorAction Stop
    
            $uploadUrl = $uploadSession.uploadUrl
            Write-Verbose "UploadUrl: $uploadUrl"
            $offset = 0
            
            while ($offset -lt $fileSize) {
                $bytesToRead = [Math]::Min($chunkSize, $fileSize - $offset)
                $buffer = New-Object byte[] $bytesToRead
                $bytesRead = $fileStream.Read($buffer, 0, $bytesToRead)
    
                if ($bytesRead -gt 0) {
                    $contentRange = "bytes $offset-$($offset + $bytesRead - 1)/$fileSize"
                    Write-Verbose "Writing range: $contentRange"
                    Invoke-GraphWithRetry `
                        -RequestUri $uploadUrl `
                        -method Put `
                        -body $buffer `
                        -headers @{ 'Content-Range' = $contentRange } `
                        -ContentType 'application/octet-stream' | out-null
                    $offset += $bytesRead
                }
            }
        }
        finally {
            $fileStream.Close()
        }
    }
}
function Get-GraphAuthorizationHeader
{
    <#
    .SYNOPSIS
    Retrieves an authorization header for Microsoft Graph API calls
    
    .DESCRIPTION
    Obtains an access token from the configured AAD authentication factory with the Graph API scope
    and returns it as a hashtable containing the Authorization header.
    This command can be called directly but is primarily used by other module functions.

    .PARAMETER FactoryName
    Optional factory name override used to obtain the token. If omitted, the factory configured
    by Set-GraphAadFactory is used.

    .INPUTS
    None
    This command does not accept pipeline input.
    
    .OUTPUTS
    System.Collections.Hashtable
    Returns a hashtable with the Authorization header containing the Bearer token.
    
    .EXAMPLE
    $authHeader = Get-GraphAuthorizationHeader
    
    Retrieves the authorization header for Graph API calls.

    .EXAMPLE
    $authHeader = Get-GraphAuthorizationHeader -FactoryName 'ManagedIdentityFactory'

    Retrieves the authorization header by explicitly selecting a token factory.
    
    .NOTES
    This function uses the scopes configured via Set-GraphScopes and the factory configured via Set-GraphAadFactory.
    #>
    param (
        $FactoryName = $script:graphConnection.FactoryName
    )

    process
    {
        Get-AadToken -Factory $FactoryName -Scope $script:graphConnection.GraphScope -AsHashTable
    }
}
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
                $result = Invoke-GraphWithRetry -RequestUri $uri -method Get -Headers $AdditionalHeaders -OperationName $OperationName -Confirm:$false -ErrorAction Stop
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
                Write-Warning "Could not retrieve data for uri: $uri. Error: $($_.Exception.Message)"
                throw
            }
        }
    }
}
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
function Invoke-GraphWithRetry
{
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    <#
    .SYNOPSIS
    Invokes a Graph API with automatic retry logic for throttling
    
    .DESCRIPTION
    Executes a Microsoft Graph API request with built-in retry logic to handle HTTP 429 (Too Many Requests) throttling responses.
    The function will automatically retry up to 100 times with exponential backoff when throttled.
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
    Maximum number of retry attempts before the error is thrown. Default is 100.

    .PARAMETER DefaultBackOffSeconds
    Fallback delay in seconds used when the response does not include Retry-After.

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
    - Automatically handles HTTP 429 throttling with exponential backoff
    - Maximum retry attempts: 100
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
        [int]$DefaultBackOffSeconds = 1
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
            $authHeader = Get-GraphAuthorizationHeader
            Write-Verbose "Invoking Graph API: $graphUri with method $method. Attempt #$($retries + 1)"
            $headers['Authorization'] = $authHeader['Authorization']
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
                    throw
                }
            }
            finally
            {
                if($null -ne $script:graphConnection.AiLogger)
                {
                    Write-AiDependency -Target 'graph.microsoft.com' -DependencyType 'Graph API' -Name $OperationName -Data $graphUri -Start $requestStart -ResultCode $resultCode -Success ($resultCode -eq 'Ok') -Connection $script:graphConnection.AiLogger
                }
            }
        }while($true)
    }
}
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
function Set-GraphAadFactory
{
    <#
    .SYNOPSIS
    Sets the AAD authentication factory for Graph API operations
    
    .DESCRIPTION
    Configures the authentication factory to be used for obtaining access tokens when making Graph API calls.
    The factory name corresponds to a factory registered with the AadAuthenticationFactory module.
    By default, the command validates that the factory exists before updating module state.
    
    .PARAMETER Name
    The name of the authentication factory to use. This should match a factory registered with AadAuthenticationFactory module.
    Common values include 'ManagedIdentityFactory' or custom factory names.

    .PARAMETER Force
    Skips validation that the specified factory exists and sets the value directly.

    .INPUTS
    None
    This command does not accept pipeline input.

    .OUTPUTS
    None
    This command updates module configuration and does not return an object.
    
    .EXAMPLE
    Set-GraphAadFactory -Name 'ManagedIdentityFactory'
    
    Configures the module to use managed identity for authentication.
    
    .EXAMPLE
    Set-GraphAadFactory -Name 'MyCustomFactory'
    
    Configures the module to use a custom authentication factory.

    .EXAMPLE
    Set-GraphAadFactory -Name 'FactoryRegisteredLater' -Force

    Sets the factory name without validating its current registration.

    .NOTES
    - When -Force is not specified, the command throws if the factory cannot be found.
    - The configured value is used by subsequent GraphApiHelper commands that request tokens.
    #>
    param
    (
        [Parameter(Mandatory)]
        [string]$Name,
        [switch]$Force
    )

    process
    {
        if($null -eq (Get-AadAuthenticationFactory -Name $Name) -and -not $Force)
        {
            throw "Authentication factory '$Name' not found. Please register it with the AadAuthenticationFactory module before using."
        }
        $script:graphConnection.FactoryName = $Name
    }
}
function Set-GraphAiLogger
{
    <#
    .SYNOPSIS
    Sets the Application Insights logger for telemetry
    
    .DESCRIPTION
    Configures the Application Insights logger instance to be used for logging telemetry data during Graph API operations.
    
    .PARAMETER Logger
    The AILogger instance to use for logging. This should be created using the ApplicationInsights module.

    .INPUTS
    None
    This command does not accept pipeline input.

    .OUTPUTS
    None
    This command updates module configuration and does not return an object.
    
    .EXAMPLE
    $aiLogger = New-AiLogger -InstrumentationKey 'your-instrumentation-key'
    Set-GraphAiLogger -Logger $aiLogger
    
    Configures the module to use the specified Application Insights logger for telemetry.

    .NOTES
    Invoke-GraphWithRetry uses this logger for dependency and exception telemetry when configured.
    #>
    param
    (
        [Parameter(Mandatory)]
        $Logger
    )

    process
    {
        $script:graphConnection.AiLogger = $Logger
    }
}
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
        $script:graphConnection.BaseUri = $BaseUri
    }
}
function Set-GraphScopes
{
    <#
    .SYNOPSIS
    Sets the scopes for Graph API authentication
    
    .DESCRIPTION
    Configures the scope to be used when requesting access tokens for Graph API calls.
    The default scope is 'https://graph.microsoft.com/.default' which uses the permissions assigned to the application in Azure AD.
    
    .PARAMETER Scopes
    The scopes to use when requesting access tokens. The default is 'https://graph.microsoft.com/.default'.

    .INPUTS
    None
    This command does not accept pipeline input.

    .OUTPUTS
    None
    This command updates module configuration and does not return an object.
    
    .EXAMPLE
    Set-GraphScopes -Scopes 'https://graph.microsoft.com/.default'
    
    Configures the module to use the default Graph API scope for authentication.
    
    .EXAMPLE
    Set-GraphScopes -Scopes 'https://graph.microsoft.com/User.Read'
    
    Configures the module to request a token with only User.Read permissions.

    .EXAMPLE
    Set-GraphScopes -Scopes @('https://graph.microsoft.com/User.Read', 'https://graph.microsoft.com/Mail.Read')

    Configures multiple delegated scopes for token acquisition.

    .NOTES
    The configured scopes are used by Get-GraphAuthorizationHeader when requesting tokens.
    #>
    param
    (
        [Parameter()]
        [string[]]$Scopes = @('https://graph.microsoft.com/.default')
    )

    process
    {
        $script:graphConnection.GraphScope = $Scopes
    }
}
#endregion Public commands
#region Internal commands
#endregion Internal commands
#region Module initialization
$script:graphConnection = [PSCustomObject]@{
    FactoryName = 'graph'
    AiLogger = $null
    BaseUri = 'https://graph.microsoft.com/v1.0'
    GraphScope = @('https://graph.microsoft.com/.default')
}
#endregion Module initialization
