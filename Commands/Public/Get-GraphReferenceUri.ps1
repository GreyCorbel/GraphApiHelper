function Get-GraphReferenceUri
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [Alias('Id')]
        [string]$ObjectId
    )

    begin
    {
        if($null -eq $script:graphConnection)
        {
            throw "Graph connection not initialized. Please call Connect-Graph first."
        }
        $script:graphConnection.GetReference($ObjectId)
    }
}