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
