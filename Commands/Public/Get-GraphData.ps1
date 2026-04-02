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
