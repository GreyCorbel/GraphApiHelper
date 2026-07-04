
function Get-GraphResponseMetadata {
    param(
        [Parameter(Mandatory)]
        [psobject]$Response
    )

    $props = $Response.PSObject.Properties

    [GraphResponseMetadata]@{
        Context  = $props['@odata.context']?.Value
        NextLink = $props['@odata.nextLink']?.Value
        DeltaLink = $props['@odata.deltaLink']?.Value
        Count    = $props['@odata.count']?.Value
        Type     = $props['@odata.type']?.Value
    }
}
