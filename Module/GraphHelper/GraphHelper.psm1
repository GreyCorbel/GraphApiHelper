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
        $graphUri = GetGraphRequestUri -Uri "$GraphFilePath"
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
    This function uses the factory configured via Set-GraphAadFactory.
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
    
    .PARAMETER OperationName
    The operation name to use for Application Insights logging. Default is 'Get-GraphData'.

    .PARAMETER AdditionalHeaders
    Additional HTTP headers to include in requests (for example ConsistencyLevel for advanced queries).
    
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
    
    .NOTES
    - Automatically handles pagination via @odata.nextLink
    - Uses Invoke-GraphWithRetry internally for throttling protection
    - Suitable for large datasets that span multiple pages
    - Uses the authentication factory configured via Set-GraphAadFactory
    #>
    param
    (
        [Parameter(Mandatory)]
        [Alias('Uri')]
        [string]$RequestUri,
        [Parameter()]
        $OperationName = 'Get-GraphData',
        [Parameter()]
        [System.Collections.Hashtable]$AdditionalHeaders = @{}
    )

    process
    {
        $uri = GetGraphRequestUri $RequestUri
        while($true)
        {
            try {
                #get page of results
                $result = Invoke-GraphWithRetry -RequestUri $uri -method Get -Headers $AdditionalHeaders -ErrorAction Stop
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
                if([string]::IsNullOrEmpty($uri))
                {
                    #no more pages
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
function Invoke-GraphWithRetry
{
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
    
    .NOTES
    - Automatically handles HTTP 429 throttling with exponential backoff
    - Maximum retry attempts: 100
    - Uses the authentication factory configured via Set-GraphAadFactory
    - Supports Application Insights telemetry when configured
    #>
    param
    (
        [Parameter(Mandatory)]
        [Alias('Uri')]
        [string]$RequestUri,
        [Parameter()]
        $method = 'Get',
        [Parameter()]
        $body,
        [Parameter()]
        $contentType = 'application/json',
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
        $graphUri = GetGraphRequestUri -Uri $RequestUri
    }
    process
    {
        do
        {
            $authHeader = Get-GraphAuthorizationHeader
            Write-Verbose "Invoking Graph API: $graphUri with method $method. Attempt #$($retries + 1)"
            $headers['Authorization'] = $authHeader['Authorization']
            $resultCode = 'Ok'
            try {
                $requestStart = Get-Date -AsUTC

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
function Set-GraphAadFactory
{
    <#
    .SYNOPSIS
    Sets the AAD authentication factory name for Graph API operations
    
    .DESCRIPTION
    Configures the authentication factory to be used for obtaining access tokens when making Graph API calls.
    The factory name corresponds to a factory registered with the AadAuthenticationFactory module.
    
    .PARAMETER Name
    The name of the authentication factory to use. This should match a factory registered with AadAuthenticationFactory module.
    Common values include 'ManagedIdentityFactory' or custom factory names.
    
    .EXAMPLE
    Set-GraphAadFactory -Name 'ManagedIdentityFactory'
    
    Configures the module to use managed identity for authentication.
    
    .EXAMPLE
    Set-GraphAadFactory -Name 'MyCustomFactory'
    
    Configures the module to use a custom authentication factory.
    #>
    param
    (
        [Parameter(Mandatory)]
        [string]$Name
    )

    process
    {
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
    
    .EXAMPLE
    $aiLogger = New-AiLogger -InstrumentationKey 'your-instrumentation-key'
    Set-GraphAiLogger -Logger $aiLogger
    
    Configures the module to use the specified Application Insights logger for telemetry.
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
    
    .EXAMPLE
    Set-GraphScopes -Scopes 'https://graph.microsoft.com/.default'
    
    Configures the module to use the default Graph API scope for authentication.
    
    .EXAMPLE
    Set-GraphScopes -Scopes 'https://graph.microsoft.com/User.Read'
    
    Configures the module to request a token with only User.Read permissions.
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
function GetGraphRequestUri
{
    param
    (
        [Parameter(Mandatory)]
        [string]$Uri
    )

    process
    {
        if(-not $uri.StartsWith('http'))
        {
            if(-not $script:graphConnection.BaseUri)
            {
                throw "BaseUri is not set. Please call Set-GraphBaseUri first or provide a full Uri"
            }
            return "$($script:graphConnection.BaseUri.TrimEnd('/'))/$($uri.TrimStart('/'))"
        }
        else
        {
            return $uri
        }
    }
}
#endregion Internal commands
#region Module initialization
$script:graphConnection = [PSCustomObject]@{
    FactoryName = 'graph'
    AiLogger = $null
    BaseUri = 'https://graph.microsoft.com/v1.0'
    GraphScope = @('https://graph.microsoft.com/.default')
}
#endregion Module initialization
