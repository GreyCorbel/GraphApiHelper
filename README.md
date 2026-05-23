# GraphApiHelper

PowerShell module for Microsoft Graph API operations with practical defaults for automation:

- Access token acquisition via AadAuthenticationFactory
- Relative or absolute Graph URI handling
- Automatic pagination for GET queries
- Retry handling for throttling (HTTP 429)
- Graph batch request helpers
- Large file upload via upload sessions
- Optional Application Insights telemetry

## Module Defaults

The module initializes with:

- Factory name: graph
- Base URI: https://graph.microsoft.com/v1.0
- Scope: https://graph.microsoft.com/.default
- AI logger: null (disabled until configured)

## Requirements

| Module | Required | Purpose |
|---|---|---|
| [AadAuthenticationFactory](https://github.com/GreyCorbel/AadAuthenticationFactory) | Yes | Provides Get-AadToken and factory registration/lookup |
| [AiLogger](https://github.com/GreyCorbel/AiLogger) | No | Optional Application Insights dependency/exception logging |

Install from PowerShell Gallery:

```powershell
Install-Module AadAuthenticationFactory
Install-Module AiLogger       # optional
Install-Module GraphApiHelper
```

## Quick Start

### 1. Configure token factory

```powershell
# Example: choose an authentication factory already registered in AadAuthenticationFactory
Set-GraphAadFactory -Name 'ManagedIdentityFactory'
```

### 2. Optional: adjust scope or endpoint

```powershell
# Default scope (app permissions)
Set-GraphScopes -Scopes 'https://graph.microsoft.com/.default'

# Sovereign cloud endpoint example
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'

# Sovereign cloud scope example
Set-GraphScopes -Scopes 'https://graph.microsoft.us/.default'
```

### 3. Read and write Graph data

```powershell
# GET with automatic pagination
$users = Get-GraphData -RequestUri '/users' -WithSelect 'id,displayName,userPrincipalName' -Top 25

# POST using retry-enabled request function
$groupBody = @{
    displayName = 'Finance Team'
    mailEnabled = $false
    mailNickname = 'finance-team'
    securityEnabled = $true
} | ConvertTo-Json

Invoke-GraphWithRetry -RequestUri '/groups' -Method Post -Body $groupBody
```

## Exported Commands

This module currently exports:

- Add-GraphLargeFile
- Add-GraphReference
- Get-GraphAuthorizationHeader
- Get-GraphData
- Invoke-GraphBatch
- Invoke-GraphWithRetry
- New-GraphBatchRequest
- New-GraphUri
- Remove-GraphReference
- Set-GraphAadFactory
- Set-GraphAiLogger
- Set-GraphBaseUri
- Set-GraphScopes

## Command Overview

### Set-GraphAadFactory

Sets the authentication factory name used when requesting tokens.

```powershell
Set-GraphAadFactory -Name 'ManagedIdentityFactory'

# Skip validation if the factory will be registered later
Set-GraphAadFactory -Name 'FactoryRegisteredLater' -Force
```

### Set-GraphScopes

Sets token scopes used by Get-GraphAuthorizationHeader and all Graph calls.

```powershell
Set-GraphScopes -Scopes 'https://graph.microsoft.com/.default'
Set-GraphScopes -Scopes @('https://graph.microsoft.com/User.Read', 'https://graph.microsoft.com/Mail.Read')
```

### Set-GraphBaseUri

Sets the base URI for relative request paths. URI must include version segment (v1.0 or beta).

```powershell
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/v1.0'
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/beta'
```

### Set-GraphAiLogger

Attaches an AiLogger connection for dependency and exception telemetry.

```powershell
$logger = Connect-AiLogger -ConnectionString '<connection-string>'
Set-GraphAiLogger -Logger $logger
```

### New-GraphUri

Builds a Graph URL from a base Uri and query options.

- Relative input + no -Relative: prepends configured BaseUri
- Absolute input + -Relative: returns path and query only
- Preserves existing query string and appends additional options

```powershell
# Absolute URL from relative path
$uri = New-GraphUri -Uri '/users' -WithSelect 'id,displayName' -WithFilter "accountEnabled eq true" -Top 25

# Relative URL for batch subrequests
$batchUrl = New-GraphUri -Uri '/users' -WithSelect 'id,displayName' -Top 5 -Relative
```

### Get-GraphAuthorizationHeader

Returns hashtable authorization header (Bearer token).

```powershell
$headers = Get-GraphAuthorizationHeader
Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/me' -Headers $headers
```

### Invoke-GraphWithRetry

Sends a single Graph request (GET/POST/PUT/PATCH/DELETE) with retry logic.

- Alias: -Uri for -RequestUri
- Default retryable code: 429
- Defaults: MaxRetries 100, DefaultBackOffSeconds 1
- Honors -WhatIf and -Confirm

```powershell
Invoke-GraphWithRetry -RequestUri '/users' -Method Get

$patch = @{ jobTitle = 'Senior Engineer' } | ConvertTo-Json
Invoke-GraphWithRetry -RequestUri '/users/john.doe@contoso.com' -Method Patch -Body $patch
```

### Get-GraphData

GET helper with automatic @odata.nextLink traversal.

- Alias: -Uri for -RequestUri
- Returns either result.value items or the single returned object
- Use -NoContinue to stop after first page
- Supports query options directly (WithSelect/WithFilter/WithCount/WithExpand/WithSearch/Top/Skip)
- Supports custom retryable HTTP status codes via -RetryableErrorCodes (default: 429)
- Supports -WhatIf and -Confirm

```powershell
# Auto-pagination
$members = Get-GraphData -RequestUri "/groups/$groupId/members"

# First page only
$firstPage = Get-GraphData -RequestUri '/users' -Top 10 -NoContinue

# Advanced query with required header
$users = Get-GraphData -RequestUri '/users' -WithFilter "startswith(displayName,'A')" -WithCount -AdditionalHeaders @{ ConsistencyLevel = 'eventual' }

# Retry both throttling and transient service errors
$users = Get-GraphData -RequestUri '/users' -RetryableErrorCodes 429,503
```

### New-GraphBatchRequest

Builds a normalized Graph batch subrequest object.

- Url must be relative (absolute URLs are rejected)
- Content-Type defaults to application/json when Body is provided

```powershell
$req1 = New-GraphBatchRequest -Id '1' -Method GET -Url '/me'
$req2 = New-GraphBatchRequest -Id '2' -Method PATCH -Url '/users/john.doe@contoso.com' -Body @{ jobTitle = 'Principal Engineer' }
```

### Invoke-GraphBatch

Posts Graph /$batch payload and returns response items.

- Accepts -BatchRequest or pipeline input
- Validates each request has id/method/url
- Enforces unique request IDs
- Enforces Graph maximum of 20 subrequests
- Supports custom retryable HTTP status codes via -RetryableErrorCodes (default: 429)
- Supports -WhatIf and -Confirm

```powershell
$requests = @(
    New-GraphBatchRequest -Id '1' -Method GET -Url '/me'
    New-GraphBatchRequest -Id '2' -Method GET -Url (New-GraphUri -Uri '/users' -Top 5 -Relative)
)

$responses = Invoke-GraphBatch -BatchRequest $requests
$responses | Select-Object id, status

# Retry outer batch call on throttling and transient service errors
$responses = Invoke-GraphBatch -BatchRequest $requests -RetryableErrorCodes 429,503
```

### Add-GraphReference / Remove-GraphReference

Adds or removes members/owners for groups, applications, or service principals.

- objectType: groups, applications, servicePrincipals
- ReferenceType: members, owners
- MemberId accepts pipeline input
- PermissiveModify suppresses already-exists (add) or not-found (remove)

```powershell
# Add members from pipeline
$memberIds | Add-GraphReference -ObjectId $groupId

# Remove owner and ignore missing reference
Remove-GraphReference -ObjectId $groupId -ReferenceType owners -MemberId $ownerId -PermissiveModify
```

### Add-GraphLargeFile

Uploads local files using Graph createUploadSession with chunked upload.

- Chunk size: 5 MB
- Conflict behavior: replace
- Uses Invoke-GraphWithRetry internally for session creation and chunk uploads

```powershell
Add-GraphLargeFile -LocalFilePath 'C:\Reports\annual-report.xlsx' -GraphFilePath '/me/drive/root:/Reports/annual-report.xlsx' -Verbose
```

## Notes

- Relative request paths are supported across the module and resolved through New-GraphUri.
- Use Invoke-GraphWithRetry for single-page or non-GET operations; use Get-GraphData for full paginated GET collection.
- This module is marked for PowerShell Core compatibility (CompatiblePSEditions = Core).

## Cloud Endpoints And Scopes

The module supports Microsoft Graph in global and national cloud instances. Configure both base URI and scope to match the target cloud.

```powershell
# Global cloud
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.com/v1.0'
Set-GraphScopes -Scopes 'https://graph.microsoft.com/.default'

# US Government (GCC High)
Set-GraphBaseUri -BaseUri 'https://graph.microsoft.us/v1.0'
Set-GraphScopes -Scopes 'https://graph.microsoft.us/.default'

# US Government DoD
Set-GraphBaseUri -BaseUri 'https://dod-graph.microsoft.us/v1.0'
Set-GraphScopes -Scopes 'https://dod-graph.microsoft.us/.default'

# China (21Vianet)
Set-GraphBaseUri -BaseUri 'https://microsoftgraph.chinacloudapi.cn/v1.0'
Set-GraphScopes -Scopes 'https://microsoftgraph.chinacloudapi.cn/.default'
```

Microsoft reference documentation:

- National cloud deployments and endpoint mapping: https://learn.microsoft.com/graph/deployments
- Microsoft identity platform scopes and permissions: https://learn.microsoft.com/entra/identity-platform/scopes-oidc
