# GraphApiHelper

PowerShell module for Microsoft Graph API operations with practical defaults for automation.

## Features

- Access token acquisition via [AadAuthenticationFactory](https://github.com/GreyCorbel/AadAuthenticationFactory)
- Relative or absolute Graph URI handling
- Automatic pagination for GET queries (`@odata.nextLink` traversal)
- Retry handling for throttling (HTTP 429) and other transient errors
- Graph batch request helpers (up to 20 subrequests per call)
- Large file upload via resumable upload sessions
- Optional Application Insights telemetry via [AiLogger](https://github.com/GreyCorbel/AiLogger)

## Requirements

| Module | Required | Purpose |
|---|---|---|
| [AadAuthenticationFactory](https://github.com/GreyCorbel/AadAuthenticationFactory) | **Yes** | Provides `Get-AadToken` and factory registration/lookup |
| [AiLogger](https://github.com/GreyCorbel/AiLogger) | No | Optional Application Insights dependency/exception logging |

## Installation

```powershell
Install-Module AadAuthenticationFactory
Install-Module AiLogger       # optional
Install-Module GraphApiHelper
```

## Module Defaults

When the module loads it initializes with these defaults:

| Setting | Default value |
|---|---|
| Factory name | `graph` |
| Base URI | `https://graph.microsoft.com/v1.0` |
| Scope | `https://graph.microsoft.com/.default` |
| AI logger | `$null` (disabled) |

All defaults can be changed with the [Configuration](Configuration) commands.

## Quick Start

### 1. Configure the authentication factory

```powershell
# Use a factory already registered in AadAuthenticationFactory
Set-GraphAadFactory -Name 'ManagedIdentityFactory'
```

### 2. (Optional) Adjust scope or endpoint

```powershell
# Delegated permissions example
Set-GraphScopes -Scopes @('User.Read', 'Mail.Read')

# Sovereign cloud
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'
Set-GraphScopes -Scopes 'https://graph.microsoft.us/.default'
```

### 3. Read and write Graph data

```powershell
# GET with automatic pagination
$users = Get-GraphData -RequestUri '/users' -WithSelect 'id,displayName,userPrincipalName' -Top 25

# POST a new group
$body = @{
    displayName    = 'Finance Team'
    mailEnabled    = $false
    mailNickname   = 'finance-team'
    securityEnabled = $true
} | ConvertTo-Json

Invoke-GraphWithRetry -RequestUri '/groups' -Method Post -Body $body
```

## Wiki Navigation

| Page | Contents |
|---|---|
| [Configuration](Configuration) | Set-GraphAadFactory, Set-GraphScopes, Set-GraphBaseUri, Set-GraphAiLogger |
| [Data Access](Data-Access) | Get-GraphData, Invoke-GraphWithRetry, Get-GraphAuthorizationHeader |
| [URL Building](URL-Building) | New-GraphUri |
| [Batch Requests](Batch-Requests) | New-GraphBatchRequest, Invoke-GraphBatch |
| [Directory References](Directory-References) | Add-GraphReference, Remove-GraphReference, Get-GraphReferenceUri |
| [File Upload](File-Upload) | Add-GraphLargeFile |
| [Error Handling](Error-Handling) | ConvertFrom-GraphErrorRecord |
| [National Clouds](National-Clouds) | Cloud endpoint and scope reference |
| [Command Reference](Command-Reference) | Quick-reference table of all exported commands |
