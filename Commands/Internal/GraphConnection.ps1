<#
.SYNOPSIS
Represents module-level connection settings for Microsoft Graph.

.DESCRIPTION
Stores shared configuration used by GraphApiHelper commands, including
the authentication factory name, Graph base URI, scopes, and optional
Application Insights logger instance.

.NOTES
This is an internal type used by module commands and is not exported.
#>
class GraphConnection {

    #name of AadAuthenticationFactry factory to use for obtaining tokens
    [string]$FactoryName
    #base URI for Microsoft Graph API calls, typically https://graph.microsoft.com/v1.0 or https://graph.microsoft.us/beta
    [Uri]$BaseUri
    #scopes required for Microsoft Graph API access
    [string[]]$GraphScope
    #optional Application Insights logger instance
    [object]$AiLogger

    GraphConnection()
    {
        #set defaults
        $this.FactoryName = 'graph'
        $this.BaseUri = [Uri]::new('https://graph.microsoft.com/v1.0')
        $this.GraphScope = @('https://graph.microsoft.com/.default')
        $this.AiLogger = $null
    }
    
    GraphConnection([string]$BaseUri, [string[]]$GraphScope, $AiLogger)
    {
        $this.FactoryName = 'graph'
        $this.BaseUri = new-object System.Uri($BaseUri)
        $this.GraphScope = $GraphScope
        $this.AiLogger = $AiLogger
    }

    <#
    .SYNOPSIS
    Builds a directory object reference URI for Microsoft Graph.

    .PARAMETER id
    The Azure AD object identifier to convert into a directoryObjects
    reference URI.

    .OUTPUTS
    System.String
    The fully-qualified reference URI for the provided object id.
    #>
    [string] GetReference([string]$id)
    {
        $ref = "$($this.BaseUri.Scheme)://$($this.BaseUri.Host)/v1.0/directoryObjects/$id"
        Write-Verbose "Constructed reference URI: $ref"
        return $ref
    }
}