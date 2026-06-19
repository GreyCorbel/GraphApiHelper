<#
.SYNOPSIS
Builds a Microsoft Graph directory object reference URI.

.DESCRIPTION
Returns the directoryObjects reference URI for the supplied object identifier
using the currently configured Graph base endpoint.

This helper is primarily used by reference-management commands such as
Add-GraphReference and Remove-GraphReference.

.PARAMETER ObjectId
The Azure AD object identifier to convert into a directoryObjects reference URI.

.INPUTS
None
This command does not accept pipeline input.

.OUTPUTS
System.String
Returns the fully-qualified Microsoft Graph reference URI.

.EXAMPLE
Get-GraphReferenceUri -ObjectId '11111111-2222-3333-4444-555555555555'

Returns a URI such as:
https://graph.microsoft.com/v1.0/directoryObjects/11111111-2222-3333-4444-555555555555

.NOTES
Throws when Graph connection state is not initialized.
#>
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