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
