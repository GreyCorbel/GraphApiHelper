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
