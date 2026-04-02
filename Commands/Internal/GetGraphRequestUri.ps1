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
