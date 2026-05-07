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
